#!/usr/bin/env python3
"""Zoorik Kubernetes cost assessment.

Reads your AKS and EKS clusters, estimates what right-sizing could save, and
writes one Excel file for you to e-mail to Zoorik. Nothing to install.

Paste this line in Azure Cloud Shell (every AKS cluster in the current
subscription), in AWS CloudShell (every EKS cluster in every enabled region),
or in any other shell, where it uses az or aws if signed in, else the current
kubectl context:
    curl -fsSL https://zoorik.com/assess/coral | python3 -

Options go after the dash. --context NAME assesses only that context:
    curl -fsSL https://zoorik.com/assess/coral | python3 - --anonymize
    curl -fsSL https://zoorik.com/assess/coral | python3 - --context prod --sample-minutes 10

The source is at github.com/Zoorikcloud/assessment.

What it does:
  - It reads. In each cluster it runs only kubectl get and kubectl top. With
    the az and aws CLIs it reads which account you are signed in to, lists
    your clusters, node pools and regions, writes each cluster's credentials
    to a temporary kubeconfig, and looks up AWS prices.
  - For a private AKS cluster kubectl cannot read from here, or an AKS
    cluster kubectl cannot reach or sign in to (kubelogin not installed), it
    uses az aks command invoke: Azure runs the same kubectl reads in a
    short-lived pod of its own, in the cluster's aks-command namespace, and
    records the call in its activity log. --no-invoke skips that; such a
    cluster is then not assessed.
  - It never reads Secrets or ConfigMaps. From what it reads it keeps only
    names, owners, labels for matching, resource requests and limits, and
    volume claim names: never env values, annotations or pod IPs.
  - It sends nothing to Zoorik. Its only calls go to your clusters, your
    cloud's own APIs and the public Azure price list (prices.azure.com),
    which receives a region and a machine type.
  - It writes one file, the Excel report. Its working files (kubectl's cache,
    and a kubeconfig for the clusters it finds through az or aws) go in a
    temporary folder that is deleted when it exits; your own kubeconfig is
    never changed.
"""

from __future__ import annotations

import argparse
import base64
import binascii
import gzip
import hashlib
import http.client
import json
import math
import os
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import time
import traceback
import urllib.error
import urllib.parse
import urllib.request
import zipfile
import zlib
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Optional
from xml.sax.saxutils import escape

# ---------------------------------------------------------------- config

VERSION = "1.0.0"
ONE_LINER = "curl -fsSL https://zoorik.com/assess/coral | python3 -"
SEND_TO = "founders@zoorik.com"

HOURS_PER_MONTH = 730
HEADROOM = 1.3
PACKING = 0.85
CPU_FLOOR = 0.010
MEM_FLOOR = 32 * 2**20
SAMPLE_INTERVAL = 30
REQUEST_TIMEOUT = "60s"
PAGE_SIZE = 500
INVOKE_TIMEOUT = 900

MIB = 2**20
GIB = 2**30
DOCS = "see github.com/Zoorikcloud/assessment#access-it-needs"
UNREACHABLE = "API server not reachable from here"
PRIVATE = "private API not reachable from here"
NETWORK = {"aks": "VNet", "eks": "VPC"}
NO_ACCESS = ("access denied", "this identity has no access to the cluster")
INVOKE_WHEN = (UNREACHABLE, "kubectl timed out", "kubelogin is not installed")
OVERHEAD_KINDS = ("DaemonSet", "Static pod")
POOL_LABELS = ("kubernetes.azure.com/agentpool", "agentpool", "eks.amazonaws.com/nodegroup", "karpenter.sh/nodepool")
AWS_ENV = dict(os.environ, AWS_PAGER="")
WORKDIR = ""


def say(text: str = "", end: str = "\n") -> None:
    print(text, end=end, flush=True)


# ---------------------------------------------------------------- commands

class Failed(Exception):
    """An expected failure, carried as a short reason."""


_REASONS = (
    ("failed to connect to msi", "Cloud Shell sign-in issue; run az login, then run this again"),
    ("executable kubelogin not found", "kubelogin is not installed"),
    ("aadsts", "Entra ID sign-in failed; run az login, then run this again"),
    ("metrics api not available", "metrics-server is not installed"),
    ("forbidden", "access denied"),
    ("authorizationfailed", "access denied"),
    ("accessdenied", "access denied"),
    ("unauthorizedoperation", "access denied"),
    ("metrics.k8s.io", "metrics-server is not installed"),
    ("unauthorized", "this identity has no access to the cluster"),
    ("must be logged in", "this identity has no access to the cluster"),
    ("az login", "not signed in to Azure; run az login"),
    ("unable to locate credentials", "no AWS credentials"),
    ("expiredtoken", "AWS credentials have expired"),
    ("i/o timeout", UNREACHABLE),
    ("dial tcp", UNREACHABLE),
    ("no such host", UNREACHABLE),
    ("connection refused", UNREACHABLE),
    ("was refused", UNREACHABLE),
    ("no route to host", UNREACHABLE),
    ("network is unreachable", UNREACHABLE),
    ("unable to connect", UNREACHABLE),
    ("deadline exceeded", UNREACHABLE),
    ("timeout exceeded", UNREACHABLE),
    ("tls handshake timeout", UNREACHABLE),
    ("timed out", UNREACHABLE),
)


def reason(text: str) -> str:
    """Turn tool error output into a short reason, free of user names."""
    low = (text or "").lower()
    for needle, short in _REASONS:
        if needle in low:
            return short
    line = next((x.strip() for x in (text or "").splitlines() if x.strip()), "unknown error")
    line = re.sub(r"^(error|ERROR|Error)[:\s]+", "", line)
    line = re.sub(r"\"[^\"]*\"|'[^']*'", "...", line)
    return line[:100]


def run(cmd: list[str], timeout: float = 120, env: Optional[dict] = None) -> str:
    """Run one command and return its output, or raise Failed.
    The Azure CLI's own usage telemetry is switched off for these calls only; your az settings stay as they are."""
    if cmd and cmd[0] == "az":
        env = dict(env or os.environ, AZURE_CORE_COLLECT_TELEMETRY="false")
    try:
        p = subprocess.run(cmd, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                           timeout=timeout, env=env, encoding="utf-8", errors="replace")
    except FileNotFoundError:
        raise Failed(f"{cmd[0]} is not installed") from None
    except subprocess.TimeoutExpired:
        raise Failed(f"{cmd[0]} timed out") from None
    if p.returncode != 0:
        raise Failed(reason(p.stderr or p.stdout))
    return p.stdout


def az(*args: str, timeout: float = 120):
    out = run(["az", *args, "-o", "json", "--only-show-errors"], timeout)
    try:
        return json.loads(out or "null")
    except ValueError:
        raise Failed("unexpected az output") from None


def aws(*args: str, region: str = "", timeout: float = 60):
    cmd = ["aws", *args, "--output", "json"] + (["--region", region] if region else [])
    out = run(cmd, timeout, env=AWS_ENV)
    try:
        return json.loads(out or "null")
    except ValueError:
        raise Failed("unexpected aws output") from None


def works(cmd: list[str]) -> bool:
    if not shutil.which(cmd[0]):
        return False
    try:
        run(cmd, 30, env=AWS_ENV)
        return True
    except Failed:
        return False


def http_json(url: str) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": f"zoorik-k8s-assessment/{VERSION}"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except (urllib.error.URLError, http.client.HTTPException, OSError, ValueError):
        raise Failed("price list not reachable") from None


# ---------------------------------------------------------------- quantities

_UNITS = {"n": 1e-9, "u": 1e-6, "m": 1e-3, "": 1.0, "k": 1e3, "K": 1e3, "M": 1e6, "G": 1e9, "T": 1e12,
          "P": 1e15, "E": 1e18, "Ki": 2.0**10, "Mi": 2.0**20, "Gi": 2.0**30, "Ti": 2.0**40, "Pi": 2.0**50,
          "Ei": 2.0**60}


_QUANTITY = re.compile(r"\s*((?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)([a-zA-Z]*)\s*")


def is_quantity(value) -> bool:
    m = _QUANTITY.fullmatch(str(value or ""))
    return bool(m) and m.group(2) in _UNITS


def quantity(value) -> float:
    """Parse a Kubernetes quantity (CPU cores or bytes); 0 when it does not parse."""
    if not is_quantity(value):
        return 0.0
    m = _QUANTITY.fullmatch(str(value))
    return float(m.group(1)) * _UNITS[m.group(2)]


# ---------------------------------------------------------------- model

@dataclass
class Pool:
    """A node pool as the cloud API reports it."""
    name: str
    instance_type: str
    count: int
    capacity: str
    os: str
    arch: str
    minimum: Optional[int]


@dataclass
class Target:
    """One cluster to assess and how to reach it."""
    name: str
    source: str = "context"
    cloud: str = ""
    region: str = ""
    version: str = ""
    context: str = ""
    kubeconfig: str = ""
    subscription: str = ""
    resource_group: str = ""
    private: bool = False
    pools: list = field(default_factory=list)
    skip: str = ""


@dataclass
class Node:
    name: str
    pool: str
    instance_type: str
    capacity: str
    os: str
    arch: str
    zone: str
    region: str
    cloud: str
    ready: bool
    cpu_cap: float
    mem_cap: float
    cpu_alloc: float
    mem_alloc: float
    pods_alloc: int
    provider: str = ""
    karpenter: bool = False


@dataclass
class Container:
    name: str
    cpu_req: float
    mem_req: float
    cpu_lim: Optional[float]
    mem_lim: Optional[float]
    sidecar: bool = False
    cpu_use: Optional[float] = None
    mem_use: Optional[float] = None

    def requests(self, recommended: bool) -> tuple:
        if not recommended:
            return self.cpu_req, self.mem_req
        return advise(self.cpu_req, self.cpu_use, CPU_FLOOR), advise(self.mem_req, self.mem_use, MEM_FLOOR)


@dataclass
class Pod:
    namespace: str
    name: str
    node: str
    labels: dict
    kind: str
    owner: str
    containers: list
    inits: list
    overhead: tuple
    claims: list
    pod_req: tuple = (None, None)
    pod_lim: tuple = (None, None)

    def request(self, recommended: bool = False) -> tuple:
        """Effective pod request as the scheduler counts it, plus overhead.

        A pod-level request wins and is kept as it is; otherwise it is
        max(running set, init phase), right-sized when recommended is set.
        """
        total = []
        for i in (0, 1):
            if self.pod_req[i] is not None:
                total.append(self.pod_req[i] + self.overhead[i])
                continue
            sidecars, init_peak = 0.0, 0.0
            for c in self.inits:
                if c.sidecar:
                    sidecars += c.requests(recommended)[i]
                    init_peak = max(init_peak, sidecars)
                else:
                    init_peak = max(init_peak, sidecars + c.requests(False)[i])
            running = sum(c.requests(recommended)[i] for c in self.containers) + sidecars
            total.append(max(running, init_peak) + self.overhead[i])
        return total[0], total[1]

    def limits(self) -> tuple:
        """Pod-level limit, else the sum over running containers; None when one has none."""
        out = []
        for i in (0, 1):
            per = [(c.cpu_lim, c.mem_lim)[i] for c in self.running()]
            own = self.pod_lim[i]
            out.append(own if own is not None else None if None in per else sum(per))
        return out[0], out[1]

    def running(self) -> list:
        return self.containers + [c for c in self.inits if c.sidecar]


@dataclass
class Workload:
    namespace: str
    kind: str
    name: str
    pods: list = field(default_factory=list)
    hpa: Optional[bool] = None
    pdb: Optional[bool] = None
    rwo: Optional[bool] = None


@dataclass
class PoolRow:
    pool: str
    instance_type: str
    capacity: str
    os: str
    arch: str
    nodes: int
    vcpu: Optional[float] = None
    mem: Optional[float] = None
    cpu_alloc: Optional[float] = None
    mem_alloc: Optional[float] = None
    ds_cpu: Optional[float] = None
    ds_mem: Optional[float] = None
    cpu_req: Optional[float] = None
    mem_req: Optional[float] = None
    cpu_rec: Optional[float] = None
    mem_rec: Optional[float] = None
    pods_max: Optional[int] = None
    pods: Optional[int] = None
    needed: Optional[int] = None
    price: Optional[float] = None
    note: str = ""

    def monthly(self, nodes: Optional[int]) -> Optional[float]:
        if self.price is None or nodes is None:
            return None
        return nodes * self.price * HOURS_PER_MONTH


@dataclass
class Cluster:
    target: Target
    cloud: str = ""
    region: str = ""
    version: str = ""
    failed: str = ""
    notes: list = field(default_factory=list)
    nodes: list = field(default_factory=list)
    pods: list = field(default_factory=list)
    workloads: list = field(default_factory=list)
    pools: list = field(default_factory=list)
    node_use: dict = field(default_factory=dict)
    ctr_use: dict = field(default_factory=dict)
    samples: int = 0
    snapshot: bool = False
    metrics_note: str = ""
    pending: int = 0


def advise(request: float, used: Optional[float], floor: float) -> float:
    """Recommended request: peak use x headroom, floored, never above today."""
    if used is None:
        return request
    recommended = max(used * HEADROOM, floor)
    return recommended if request == 0 else min(recommended, request)


# ---------------------------------------------------------------- kubectl

NODE_JQ = ('[.items[]|{metadata:{name:.metadata.name,labels:((.metadata.labels//{})|with_entries(select(.key|'
           'test("instance-type|agentpool|nodegroup|nodepool|capacity|scalesetpriority|topology|kubernetes.io/os|'
           'kubernetes.io/arch"))))},spec:{providerID:.spec.providerID},status:{allocatable:.status.allocatable,'
           'capacity:.status.capacity,conditions:[.status.conditions[]?|select(.type=="Ready")|{type,status}]}}]')
POD_JQ = ('[.items[]|select(.status.phase!="Succeeded" and .status.phase!="Failed")|{metadata:{name:.metadata.name,'
          'namespace:.metadata.namespace,labels:.metadata.labels,ownerReferences:[.metadata.ownerReferences[]?|'
          '{kind,name,controller}]},spec:{nodeName:.spec.nodeName,overhead:.spec.overhead,resources:.spec.resources,'
          'containers:[.spec.containers[]|{name,resources}],initContainers:[.spec.initContainers[]?|{name,resources,'
          'restartPolicy}],volumes:[.spec.volumes[]?|select(.persistentVolumeClaim)|{persistentVolumeClaim:'
          '{claimName:.persistentVolumeClaim.claimName}}]},status:{phase:.status.phase}}]')
OWNED_JQ = ('[.items[]|select(.metadata.ownerReferences)|{metadata:{name:.metadata.name,'
            'namespace:.metadata.namespace,ownerReferences:[.metadata.ownerReferences[]|{kind,name,controller}]}}]')
HPA_JQ = '[.items[]|{metadata:{namespace:.metadata.namespace},spec:{scaleTargetRef:.spec.scaleTargetRef}}]'
PDB_JQ = '[.items[]|{metadata:{namespace:.metadata.namespace},spec:{selector:.spec.selector}}]'
PVC_JQ = ('[.items[]|{metadata:{name:.metadata.name,namespace:.metadata.namespace},'
          'spec:{accessModes:.spec.accessModes}}]')

RESOURCES = {
    "nodes": ("nodes", "/api/v1/nodes", NODE_JQ),
    "pods": ("pods", "/api/v1/pods", POD_JQ),
    "replicasets": ("replicasets.apps", "/apis/apps/v1/replicasets", OWNED_JQ),
    "jobs": ("jobs.batch", "/apis/batch/v1/jobs", OWNED_JQ),
    "hpa": ("horizontalpodautoscalers.autoscaling", "/apis/autoscaling/v1/horizontalpodautoscalers", HPA_JQ),
    "pdb": ("poddisruptionbudgets.policy", "/apis/policy/v1/poddisruptionbudgets", PDB_JQ),
    "pvc": ("persistentvolumeclaims", "/api/v1/persistentvolumeclaims", PVC_JQ),
}
OPTIONAL = {"replicasets": "ReplicaSets", "jobs": "Jobs", "hpa": "HPAs", "pdb": "PDBs", "pvc": "volume claims"}


def slim_pod(o: dict) -> Optional[dict]:
    """Keep what POD_JQ keeps; drop finished pods, env values and everything else."""
    md, spec, phase = o.get("metadata") or {}, o.get("spec") or {}, (o.get("status") or {}).get("phase")
    if phase in ("Succeeded", "Failed"):
        return None

    def keep(items, *keys):
        return [{k: x.get(k) for k in keys} for x in items or []]

    return {"metadata": {k: md.get(k) for k in ("name", "namespace", "labels", "ownerReferences")},
            "spec": {"nodeName": spec.get("nodeName"), "overhead": spec.get("overhead"),
                     "resources": spec.get("resources"),
                     "containers": keep(spec.get("containers"), "name", "resources"),
                     "initContainers": keep(spec.get("initContainers"), "name", "resources", "restartPolicy"),
                     "volumes": [v for v in spec.get("volumes") or [] if v.get("persistentVolumeClaim")]},
            "status": {"phase": phase}}


def slim_owned(o: dict) -> Optional[dict]:
    """Keep what OWNED_JQ keeps: name, namespace and owners; drop the pod template."""
    md = o.get("metadata") or {}
    if not md.get("ownerReferences"):
        return None
    return {"metadata": {k: md.get(k) for k in ("name", "namespace", "ownerReferences")}}


SLIM = {"pods": slim_pod, "replicasets": slim_owned, "jobs": slim_owned}


def kubectl(t: Target, *args: str, timeout: float = 300, request_timeout: str = REQUEST_TIMEOUT) -> str:
    cmd = ["kubectl", "--cache-dir", os.path.join(WORKDIR, "cache"), "--request-timeout", request_timeout]
    if t.kubeconfig:
        cmd += ["--kubeconfig", t.kubeconfig]
    if t.context:
        cmd += ["--context", t.context]
    return run(cmd + list(args), timeout)


def list_paged(t: Target, kind: str) -> list:
    """List one kind page by page, keeping only the fields the report needs."""
    path, slim = RESOURCES[kind][1], SLIM.get(kind)
    items, token = [], ""
    while True:
        query = {"limit": PAGE_SIZE, **({"continue": token} if token else {})}
        out = kubectl(t, "get", "--raw", f"{path}?{urllib.parse.urlencode(query)}", timeout=120)
        try:
            page = json.loads(out)
        except ValueError:
            raise Failed("unexpected kubectl output") from None
        page_items = page.get("items") or []
        items += [x for x in map(slim, page_items) if x] if slim else page_items
        token = (page.get("metadata") or {}).get("continue")
        if not token:
            return items


def fetch_direct(t: Target) -> dict:
    """Read every resource kind with kubectl get."""
    try:
        answer = kubectl(t, "get", "--raw", "/version", timeout=60, request_timeout="20s")
        version = json.loads(answer).get("gitVersion", "")
    except (ValueError, AttributeError):
        raise Failed("unexpected /version answer") from None
    snap = {"version": version, "errors": {}}
    with ThreadPoolExecutor(4) as ex:
        futures = {kind: ex.submit(list_paged, t, kind) for kind in RESOURCES}
    for kind, fut in futures.items():
        try:
            snap[kind] = fut.result()
        except Failed as e:
            snap[kind], snap["errors"][kind] = [], str(e)
    return snap


def top_direct(t: Target) -> tuple:
    nodes = parse_top_nodes(kubectl(t, "top", "nodes", "--no-headers", timeout=120))
    pods = parse_top_pods(kubectl(t, "top", "pods", "-A", "--containers", "--no-headers", timeout=180))
    if not nodes:
        raise Failed("metrics-server returned no data")
    return nodes, pods


def parse_top_nodes(text: str) -> dict:
    """NAME CPU CPU% MEMORY MEMORY%; lines that do not parse are skipped."""
    check_top(text)
    out = {}
    for line in text.splitlines():
        cols = line.split()
        if len(cols) >= 4 and is_quantity(cols[1]) and is_quantity(cols[3]):
            out[cols[0]] = (quantity(cols[1]), quantity(cols[3]))
    return out


def parse_top_pods(text: str) -> dict:
    """NAMESPACE POD CONTAINER CPU MEMORY; lines that do not parse are skipped."""
    check_top(text)
    out = {}
    for line in text.splitlines():
        cols = line.split()
        if len(cols) >= 5 and is_quantity(cols[3]) and is_quantity(cols[4]):
            out[(cols[0], cols[1], cols[2])] = (quantity(cols[3]), quantity(cols[4]))
    return out


def check_top(text: str) -> None:
    for line in text.splitlines():
        line = line.strip()
        if line.startswith("!"):
            raise Failed(reason(line[1:]))
        if line.lower().startswith("error"):
            raise Failed(reason(line))


# ---------------------------------------------------------------- command invoke

_HELPERS = (
    "E=$(mktemp 2>/dev/null || echo /dev/null); "
    "r(){ o=$(\"$@\" 2>\"$E\") || { printf '!%s\\n' \"$(head -c 300 \"$E\" | tr '\\n' ' ')\"; return 1; }; }; "
    "p(){ r \"$@\" && printf '%s\\n' \"$o\"; }; "
    "g(){ r kubectl get \"$1\" -A -o json --request-timeout=60s --chunk-size=500 && "
    "{ printf '%s' \"$o\" | jq -c \"$2\" 2>/dev/null || echo '!unreadable output'; }; }")
_SECTIONS = {
    "version": "p kubectl get --raw /version --request-timeout=60s",
    "top_nodes": "p kubectl top nodes --no-headers --request-timeout=60s",
    "top_pods": "p kubectl top pods -A --containers --no-headers --request-timeout=60s",
}
_SECTIONS.update({kind: f"g {res} '{jq}'" for kind, (res, _, jq) in RESOURCES.items()})


class Truncated(Exception):
    """The invoke output hit Azure's size limit."""


def bundle_script(sections: list) -> str:
    """One shell line that prints the sections, gzip+base64 when available.

    kubectl's stderr never enters the data: a failed command prints one line
    starting with '!' and the start of its error instead.
    """
    body = "; ".join(f"echo @@{name}; {_SECTIONS[name]}" for name in sections)
    return (f"{_HELPERS}; b(){{ {body}; echo @@end; }}; "
            "if command -v gzip >/dev/null 2>&1 && command -v base64 >/dev/null 2>&1; "
            "then echo @@gz; b | gzip -c | base64; echo @@gzend; else b; fi; "
            "[ \"$E\" = /dev/null ] || rm -f \"$E\"")


def parse_bundle(logs: str) -> dict:
    if "@@gz\n" in logs:
        body = logs.split("@@gz\n", 1)[1]
        if "@@gzend" not in body:
            raise Truncated()
        try:
            logs = gzip.decompress(base64.b64decode(body.split("@@gzend", 1)[0])).decode("utf-8", "replace")
        except (binascii.Error, OSError, EOFError, zlib.error):
            raise Truncated() from None
    sections, name = {}, None
    for line in logs.splitlines():
        if line.startswith("@@"):
            name = line[2:].strip()
            sections[name] = []
        elif name:
            sections[name].append(line)
    if "end" not in sections:
        raise Truncated()
    return {k: "\n".join(v) for k, v in sections.items() if k != "end"}


def az_invoke(t: Target, script: str) -> str:
    out = run(["az", "aks", "command", "invoke", "--subscription", t.subscription, "-g", t.resource_group,
               "-n", t.name, "--command", script, "-o", "json", "--only-show-errors"], INVOKE_TIMEOUT)
    try:
        data = json.loads(out)
    except ValueError:
        raise Failed("unexpected command invoke output") from None
    if data.get("provisioningState") not in (None, "Succeeded"):
        raise Failed(reason(data.get("reason") or "command invoke failed"))
    return data.get("logs") or ""


def fetch_invoke(t: Target) -> dict:
    """Read the cluster through az aks command invoke: one call, or one per kind if too large."""
    names = ["version", *RESOURCES, "top_nodes", "top_pods"]
    failed = {}
    try:
        sections = parse_bundle(az_invoke(t, bundle_script(names)))
    except Truncated:
        say("output over Azure's 512 KB limit, reading kind by kind ... ", end="")
        sections = {}
        for name in names:
            try:
                sections.update(parse_bundle(az_invoke(t, bundle_script([name]))))
            except Truncated:
                failed[name] = "too large for command invoke (512 KB limit)"
            except Failed as e:
                failed[name] = str(e)
    texts = {}
    for name in names:
        text = sections.get(name, "").strip()
        if name in failed:
            continue
        if text.startswith("!"):
            failed[name] = reason(text[1:])
        elif not text:
            failed[name] = "no output"
        else:
            texts[name] = text
    if "version" in failed:
        raise Failed(failed["version"])
    try:
        version = json.loads(texts["version"]).get("gitVersion", "")
    except (ValueError, AttributeError):
        raise Failed("unexpected /version answer") from None
    snap = {"version": version, "errors": {}, "top_nodes": texts.get("top_nodes", ""),
            "top_pods": texts.get("top_pods", ""), "top_error": failed.get("top_nodes") or failed.get("top_pods")}
    for kind in RESOURCES:
        if kind in failed:
            snap[kind], snap["errors"][kind] = [], failed[kind]
            continue
        try:
            snap[kind] = json.loads(texts[kind])
        except ValueError:
            snap[kind], snap["errors"][kind] = [], "unreadable output"
    snap["pods"] = [p for p in snap["pods"] if (p.get("metadata") or {}).get("namespace") != "aks-command"]
    return snap


# ---------------------------------------------------------------- discovery

def in_azure_cloud_shell() -> bool:
    return bool(os.environ.get("ACC_CLOUD")) or os.environ.get("AZUREPS_HOST_ENVIRONMENT", "").startswith("cloud-shell")


def in_aws_cloudshell() -> bool:
    return os.environ.get("AWS_EXECUTION_ENV", "") == "CloudShell"


def discover_azure(subscription: str) -> list:
    acct = az("account", "show", *(["--subscription", subscription] if subscription else []))
    sub = acct.get("id", "")
    targets = []
    for c in az("aks", "list", "--subscription", sub, timeout=180) or []:
        access = c.get("apiServerAccessProfile") or {}
        t = Target(name=c.get("name", ""), source="aks", cloud="azure", region=c.get("location", ""),
                   version=c.get("currentKubernetesVersion") or c.get("kubernetesVersion") or "",
                   subscription=sub, resource_group=c.get("resourceGroup", ""),
                   private=bool(access.get("enablePrivateCluster")))
        for p in c.get("agentPoolProfiles") or []:
            t.pools.append(Pool(name=p.get("name", ""), instance_type=p.get("vmSize", ""), count=p.get("count") or 0,
                                capacity="spot" if p.get("scaleSetPriority") == "Spot" else "on-demand",
                                os=(p.get("osType") or "linux").lower(), arch="",
                                minimum=p.get("minCount") if p.get("enableAutoScaling") else None))
        if (c.get("powerState") or {}).get("code") == "Stopped":
            t.skip = "cluster is stopped"
        targets.append(t)
    say(f"Azure subscription {acct.get('name', '')} ({sub}): {len(targets)} AKS cluster(s)")
    return targets


def discover_aws(regions: list) -> list:
    ident = aws("sts", "get-caller-identity")
    if not regions:
        home = os.environ.get("AWS_REGION") or os.environ.get("AWS_DEFAULT_REGION") or "us-east-1"
        try:
            regions = aws("ec2", "describe-regions", "--query", "Regions[].RegionName", region=home) or [home]
        except Failed as e:
            say(f"AWS: could not list regions ({e}); searching {home} only (add --region NAME for others)")
            regions = [home]
    errors = {}

    def in_region(region: str) -> list:
        try:
            return [(region, n) for n in (aws("eks", "list-clusters", region=region) or {}).get("clusters", [])]
        except Failed as e:
            errors[region] = str(e)
            return []

    with ThreadPoolExecutor(8) as ex:
        found = [x for xs in ex.map(in_region, sorted(regions)) for x in xs]
        if len(errors) == len(regions):
            raise Failed(errors[min(errors)])
        targets = list(ex.map(lambda rn: describe_eks(*rn), found))
    say(f"AWS account {ident.get('Account', '')}: {len(targets)} EKS cluster(s) in "
        f"{len(regions) - len(errors)} region(s)")
    if errors:
        say(f"  could not search {', '.join(sorted(errors))}: {'; '.join(sorted(set(errors.values())))}")
    return targets


def describe_eks(region: str, name: str) -> Target:
    t = Target(name=name, source="eks", cloud="aws", region=region)
    try:
        c = aws("eks", "describe-cluster", "--name", name, region=region)["cluster"]
    except (Failed, KeyError, TypeError) as e:
        t.skip = str(e) if isinstance(e, Failed) else "unexpected describe-cluster output"
        return t
    t.version = c.get("version", "")
    t.private = not (c.get("resourcesVpcConfig") or {}).get("endpointPublicAccess", True)
    if c.get("status") != "ACTIVE":
        t.skip = f"cluster status is {c.get('status')}"
    try:
        for g in (aws("eks", "list-nodegroups", "--cluster-name", name, region=region) or {}).get("nodegroups", []):
            ng = aws("eks", "describe-nodegroup", "--cluster-name", name, "--nodegroup-name", g,
                     region=region)["nodegroup"]
            sc, ami = ng.get("scalingConfig") or {}, ng.get("amiType", "")
            custom = ami == "CUSTOM"
            t.pools.append(Pool(name=g, instance_type=",".join(ng.get("instanceTypes") or []),
                                count=sc.get("desiredSize") or 0,
                                capacity={"SPOT": "spot", "ON_DEMAND": "on-demand"}.get(ng.get("capacityType"),
                                                                                       "unknown"),
                                os="" if custom else "windows" if "WINDOWS" in ami else "linux",
                                arch="" if custom else "arm64" if "ARM" in ami else "amd64",
                                minimum=sc.get("minSize")))
    except (Failed, KeyError, TypeError):
        t.pools = []
    return t


def credentials(t: Target) -> None:
    """Write this cluster's credentials to a temporary kubeconfig."""
    path = os.path.join(WORKDIR, re.sub(r"[^A-Za-z0-9_.-]", "_", f"{t.source}-{t.region}-{t.resource_group}-{t.name}"))
    if t.source == "aks":
        try:
            run(["az", "aks", "get-credentials", "--subscription", t.subscription, "-g", t.resource_group,
                 "-n", t.name, "--file", path, "--overwrite-existing", "--only-show-errors"], 120)
        except Failed as e:
            hint = f" (it needs the Azure Kubernetes Service Cluster User Role; {DOCS})" if str(e) in NO_ACCESS else ""
            raise Failed(f"{e}{hint}") from None
        with open(path, encoding="utf-8") as f:
            needs_kubelogin = "kubelogin" in f.read()
        if needs_kubelogin:
            if not shutil.which("kubelogin"):
                raise Failed("kubelogin is not installed")
            run(["kubelogin", "convert-kubeconfig", "-l", "azurecli", "--kubeconfig", path], 60)
    else:
        run(["aws", "eks", "update-kubeconfig", "--name", t.name, "--region", t.region, "--kubeconfig", path],
            120, env=AWS_ENV)
    t.kubeconfig = path


def kube_contexts(all_contexts: bool) -> list:
    if all_contexts:
        names = run(["kubectl", "config", "get-contexts", "-o", "name"], 30).split()
    else:
        names = [run(["kubectl", "config", "current-context"], 30).strip()]
    if not any(names):
        raise Failed("no kube context is set")
    return [Target(name=n, context=n) for n in names if n]


def choose_targets(args, shell: str) -> tuple:
    """Return (targets, whether a cloud search failed, why kubectl found none) from the flags or where it runs."""
    discover = list(args.discover or [])
    if args.subscription and "azure" not in discover:
        discover.append("azure")
    if args.region and "aws" not in discover:
        discover.append("aws")
    explicit = bool(args.context or args.all_contexts or discover)
    if not explicit:
        if shell == "azure" or (not shell and works(["az", "account", "show", "-o", "none"])):
            discover.append("azure")
        if shell == "aws" or (not shell and works(["aws", "sts", "get-caller-identity"])):
            discover.append("aws")
    targets, search_failed, why = [], False, ""
    for cloud in discover:
        try:
            targets += discover_azure(args.subscription or "") if cloud == "azure" else discover_aws(args.region or [])
        except Failed as e:
            search_failed = True
            say(f"{'Azure' if cloud == 'azure' else 'AWS'}: could not list clusters: {e}")
    try:
        if args.all_contexts:
            targets += kube_contexts(True)
        targets += [Target(name=c, context=c) for c in args.context or []]
        if not explicit and not targets:
            current = kube_contexts(False)
            kubectl(current[0], "get", "--raw", "/version", timeout=60, request_timeout="20s")
            targets = current
    except Failed as e:
        why = f"kubectl: {e}"
    seen, unique = set(), []
    for t in targets:
        key = (t.source, t.context, t.subscription, t.resource_group, t.region, t.name)
        if key not in seen:
            seen.add(key)
            unique.append(t)
    return unique, search_failed, why


# ---------------------------------------------------------------- collectors

def collect(t: Target, allow_invoke: bool) -> Cluster:
    """Inventory one cluster; record the reason and go on when it fails."""
    c = Cluster(target=t, cloud=t.cloud, region=t.region, version=t.version)
    if t.skip:
        c.failed = t.skip
        return c
    try:
        if t.source in ("aks", "eks"):
            credentials(t)
        build(c, fetch_direct(t))
        return c
    except Failed as e:
        error = str(e)
    if t.private and error in (UNREACHABLE, "kubectl timed out"):
        error = PRIVATE
    if t.source == "aks" and allow_invoke and (t.private or error in INVOKE_WHEN):
        try:
            say("via az aks command invoke (about a minute) ... ", end="")
            snap = fetch_invoke(t)
            build(c, snap)
            c.snapshot = True
            try:
                if snap["top_error"]:
                    raise Failed(snap["top_error"])
                c.node_use = parse_top_nodes(snap["top_nodes"])
                c.ctr_use = parse_top_pods(snap["top_pods"])
                c.samples = 1 if c.node_use else 0
                c.metrics_note = "" if c.node_use else "metrics-server returned no data"
            except Failed as e:
                c.node_use, c.ctr_use, c.metrics_note = {}, {}, str(e)
            return c
        except Failed as e:
            c.failed = f"{error}; command invoke: {e}"
    elif error == PRIVATE and t.source in NETWORK:
        c.failed = f"{PRIVATE}; run from a machine inside the {NETWORK[t.source]}"
    elif error in NO_ACCESS:
        c.failed = f"{error} ({'it needs an EKS access entry; ' if t.source == 'eks' else ''}{DOCS})"
    else:
        c.failed = error
    return c


def build(c: Cluster, snap: dict) -> None:
    """Turn raw objects into nodes, pods and workloads."""
    errors = snap["errors"]
    for kind in ("nodes", "pods"):
        if kind in errors:
            denied = errors[kind] in NO_ACCESS
            hint = f" (it needs cluster-wide read access to nodes and pods; {DOCS})" if denied else ""
            raise Failed(f"cannot list {kind}: {errors[kind]}{hint}")
    for kind, label in OPTIONAL.items():
        if kind in errors:
            c.notes.append(f"{label} not readable: {errors[kind]}")
    c.version = snap["version"] or c.version
    c.nodes = [make_node(o) for o in snap["nodes"]]
    hints = Counter(n.cloud for n in c.nodes if n.cloud)
    c.cloud = hints.most_common(1)[0][0] if hints else c.cloud or "other"
    regions = Counter(n.region for n in c.nodes if n.region)
    c.region = regions.most_common(1)[0][0] if regions else c.region

    owners = {}
    for kind, items in (("ReplicaSet", snap["replicasets"]), ("Job", snap["jobs"])):
        for o in items:
            md = o.get("metadata") or {}
            ref = controller_ref(md)
            if ref:
                owners[(kind, md.get("namespace"), md.get("name"))] = (ref.get("kind", ""), ref.get("name", ""))
    known_nodes = {n.name for n in c.nodes}
    for o in snap["pods"]:
        if (o.get("status") or {}).get("phase") in ("Succeeded", "Failed"):
            continue
        pod = make_pod(o, owners)
        if pod.node in known_nodes:
            c.pods.append(pod)
        elif not pod.node:
            c.pending += 1
    if c.pending:
        c.notes.append(f"{c.pending} pending pod(s) not placed in the estimate")

    hpa = None if "hpa" in errors else {
        ((o.get("metadata") or {}).get("namespace"), ref.get("kind"), ref.get("name"))
        for o in snap["hpa"] for ref in [(o.get("spec") or {}).get("scaleTargetRef") or {}]}
    pdbs = None if "pdb" in errors else [
        ((o.get("metadata") or {}).get("namespace"), (o.get("spec") or {}).get("selector")) for o in snap["pdb"]]
    rwo = None if "pvc" in errors else {
        ((o.get("metadata") or {}).get("namespace"), (o.get("metadata") or {}).get("name"))
        for o in snap["pvc"]
        if {"ReadWriteOnce", "ReadWriteOncePod"} & set((o.get("spec") or {}).get("accessModes") or [])}

    groups = {}
    for p in c.pods:
        w = groups.setdefault((p.namespace, p.kind, p.owner), Workload(p.namespace, p.kind, p.owner))
        w.pods.append(p)
    for w in groups.values():
        w.hpa = None if hpa is None else (w.namespace, w.kind, w.name) in hpa
        w.pdb = None if pdbs is None else any(
            ns == w.namespace and selects(sel, p.labels) for p in w.pods for ns, sel in pdbs)
        w.rwo = None if rwo is None else any((w.namespace, claim) in rwo for p in w.pods for claim in p.claims)
    c.workloads = sorted(groups.values(), key=lambda w: (w.namespace, w.kind, w.name))


def make_node(o: dict) -> Node:
    md, st = o.get("metadata") or {}, o.get("status") or {}
    labels = md.get("labels") or {}
    alloc, cap = st.get("allocatable") or {}, st.get("capacity") or {}
    provider = (o.get("spec") or {}).get("providerID") or ""
    if provider.startswith("azure://") or any(k.startswith("kubernetes.azure.com/") for k in labels):
        cloud = "azure"
    elif provider.startswith("aws://") or any(k.startswith("eks.amazonaws.com/") for k in labels):
        cloud = "aws"
    else:
        cloud = ""
    aws_capacity = (labels.get("karpenter.sh/capacity-type") or labels.get("eks.amazonaws.com/capacityType") or "")
    aws_capacity = aws_capacity.lower().replace("_", "-")
    if labels.get("kubernetes.azure.com/scalesetpriority", "").lower() == "spot":
        capacity = "spot"
    elif aws_capacity in ("spot", "on-demand"):
        capacity = aws_capacity
    elif any(k.startswith("kubernetes.azure.com/") for k in labels):
        capacity = "on-demand"
    else:
        capacity = "unknown"
    return Node(
        name=md.get("name", ""),
        pool=next((labels[k] for k in POOL_LABELS if labels.get(k)), "unlabelled"),
        instance_type=(labels.get("node.kubernetes.io/instance-type")
                       or labels.get("beta.kubernetes.io/instance-type", "")),
        capacity=capacity,
        os=labels.get("kubernetes.io/os", ""),
        arch=labels.get("kubernetes.io/arch", ""),
        zone=labels.get("topology.kubernetes.io/zone", ""),
        region=labels.get("topology.kubernetes.io/region", ""),
        cloud=cloud,
        ready=any(x.get("type") == "Ready" and x.get("status") == "True" for x in st.get("conditions") or []),
        cpu_cap=quantity(cap.get("cpu")), mem_cap=quantity(cap.get("memory")),
        cpu_alloc=quantity(alloc.get("cpu")), mem_alloc=quantity(alloc.get("memory")),
        pods_alloc=int(quantity(alloc.get("pods"))),
        provider=provider,
        karpenter="karpenter.sh/nodepool" in labels)


def controller_ref(md: dict) -> Optional[dict]:
    refs = md.get("ownerReferences") or []
    return next((r for r in refs if r.get("controller")), refs[0] if refs else None)


def make_container(o: dict, sidecar: bool = False) -> Container:
    res = o.get("resources") or {}
    req, lim = res.get("requests") or {}, res.get("limits") or {}
    return Container(name=o.get("name", ""),
                     cpu_req=quantity(req.get("cpu", lim.get("cpu"))),
                     mem_req=quantity(req.get("memory", lim.get("memory"))),
                     cpu_lim=quantity(lim["cpu"]) if "cpu" in lim else None,
                     mem_lim=quantity(lim["memory"]) if "memory" in lim else None,
                     sidecar=sidecar)


def make_pod(o: dict, owners: dict) -> Pod:
    md, spec = o.get("metadata") or {}, o.get("spec") or {}
    ns, name, node = md.get("namespace", ""), md.get("name", ""), spec.get("nodeName") or ""
    labels = md.get("labels") or {}
    template_hash = labels.get("pod-template-hash", "")
    kind, owner = "Pod", name
    ref = controller_ref(md)
    if ref:
        kind, owner = ref.get("kind", ""), ref.get("name", "")
        if (kind, ns, owner) in owners:
            kind, owner = owners[(kind, ns, owner)]
        elif kind == "ReplicaSet" and template_hash and owner.endswith("-" + template_hash):
            kind, owner = "Deployment", owner[:-len(template_hash) - 1]
        elif kind == "Node":
            kind, owner = "Static pod", name[:-len(node) - 1] if node and name.endswith("-" + node) else name
    overhead = spec.get("overhead") or {}
    own = spec.get("resources") or {}
    own_req, own_lim = own.get("requests") or {}, own.get("limits") or {}
    return Pod(namespace=ns, name=name, node=node, labels=labels, kind=kind, owner=owner,
               containers=[make_container(x) for x in spec.get("containers") or []],
               inits=[make_container(x, x.get("restartPolicy") == "Always") for x in spec.get("initContainers") or []],
               overhead=(quantity(overhead.get("cpu")), quantity(overhead.get("memory"))),
               claims=[(v.get("persistentVolumeClaim") or {}).get("claimName", "") for v in spec.get("volumes") or []
                       if v.get("persistentVolumeClaim")],
               pod_req=tuple(quantity(own_req[r]) if r in own_req else None for r in ("cpu", "memory")),
               pod_lim=tuple(quantity(own_lim[r]) if r in own_lim else None for r in ("cpu", "memory")))


def selects(selector: Optional[dict], labels: dict) -> bool:
    """Kubernetes label selector match; an empty selector matches every pod."""
    if selector is None:
        return False
    for k, v in (selector.get("matchLabels") or {}).items():
        if labels.get(k) != v:
            return False
    for e in selector.get("matchExpressions") or []:
        key, op, values = e.get("key"), e.get("operator"), e.get("values") or []
        if ((op == "In" and labels.get(key) not in values)
                or (op == "NotIn" and key in labels and labels[key] in values)
                or (op == "Exists" and key not in labels)
                or (op == "DoesNotExist" and key in labels)):
            return False
    return True


def sample(c: Cluster) -> None:
    """Take one kubectl top sample and keep the peak per node and container."""
    try:
        nodes, ctrs = top_direct(c.target)
    except Failed as e:
        if not c.samples:
            c.metrics_note = str(e)
        return
    for store, fresh in ((c.node_use, nodes), (c.ctr_use, ctrs)):
        for key, (cpu, mem) in fresh.items():
            old = store.get(key, (0.0, 0.0))
            store[key] = (max(old[0], cpu), max(old[1], mem))
    c.samples += 1
    c.metrics_note = ""


def sample_all(clusters: list, minutes: float) -> None:
    live = [c for c in clusters if not c.failed and not c.snapshot]
    if not live:
        return
    rounds = int(minutes * 60 // SAMPLE_INTERVAL) + 1
    if rounds > 1:
        say(f"Sampling usage for {minutes:g} min ({rounds} samples per cluster) ...")
    start = time.monotonic()
    with ThreadPoolExecutor(min(8, len(live))) as ex:
        for i in range(rounds):
            due = [c for c in live if c.samples or i == 0]
            if not due:
                return
            time.sleep(max(0.0, start + i * SAMPLE_INTERVAL - time.monotonic()))
            list(ex.map(sample, due))


def apply_usage(c: Cluster) -> None:
    for p in c.pods:
        for ctr in p.running():
            use = c.ctr_use.get((p.namespace, p.name, ctr.name))
            if use:
                ctr.cpu_use, ctr.mem_use = use


# ---------------------------------------------------------------- pricing

class Pricer:
    """Hourly list prices per machine type, cached, never guessed."""

    def __init__(self, enabled: bool):
        self.enabled = enabled
        self.cache = {}

    def hourly(self, cloud: str, region: str, sku: str, capacity: str, os_name: str) -> tuple:
        if not self.enabled:
            return None, "pricing skipped (--no-prices)"
        if cloud not in ("azure", "aws"):
            return None, "no public price list for this platform"
        if not sku:
            return None, "instance type unknown"
        if "," in sku:
            return None, "several instance types in one node group"
        if capacity not in ("spot", "on-demand"):
            return None, "Spot or on-demand not known from the node labels"
        if os_name and os_name != "linux":
            return None, f"{os_name} nodes are not priced"
        if not re.fullmatch(r"[A-Za-z0-9_.-]+", region or "") or not re.fullmatch(r"[A-Za-z0-9_.-]+", sku):
            return None, "region unknown"
        key = (cloud, region, sku, capacity)
        if key not in self.cache:
            try:
                fetch = azure_price if cloud == "azure" else aws_price
                self.cache[key] = fetch(region, sku, capacity == "spot")
            except Failed as e:
                self.cache[key] = (None, str(e) if str(e).startswith("no ") else f"price lookup: {e}")
            except (ValueError, KeyError, TypeError, AttributeError):
                self.cache[key] = (None, "price lookup failed")
        return self.cache[key]


def azure_price(region: str, sku: str, spot: bool) -> tuple:
    query = (f"armRegionName eq '{region}' and armSkuName eq '{sku}' "
             "and serviceName eq 'Virtual Machines' and priceType eq 'Consumption'")
    url = "https://prices.azure.com/api/retail/prices?" + urllib.parse.urlencode({"$filter": query})
    prices = []
    while url:
        data = http_json(url)
        for item in data.get("Items") or []:
            meter, product = item.get("meterName", ""), item.get("productName", "")
            if "Windows" in product or "Low Priority" in meter or item.get("unitOfMeasure") != "1 Hour":
                continue
            if meter.endswith("Spot") == spot and (item.get("unitPrice") or 0) > 0:
                prices.append(float(item["unitPrice"]))
        url = data.get("NextPageLink")
    if not prices:
        raise Failed(f"no list price for {sku} in {region}")
    return min(prices), "Azure retail price, Linux " + ("Spot" if spot else "pay-as-you-go")


def aws_price(region: str, itype: str, spot: bool) -> tuple:
    if spot:
        now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        out = aws("ec2", "describe-spot-price-history", "--instance-types", itype,
                  "--product-descriptions", "Linux/UNIX", "--start-time", now, region=region) or {}
        latest = {}
        for h in sorted(out.get("SpotPriceHistory") or [], key=lambda h: h.get("Timestamp", "")):
            latest[h.get("AvailabilityZone")] = float(h.get("SpotPrice") or 0)
        if not latest:
            raise Failed(f"no Spot price for {itype} in {region}")
        return sum(latest.values()) / len(latest), f"AWS Spot price, average of {len(latest)} zone(s)"
    terms = (("instanceType", itype), ("regionCode", region), ("operatingSystem", "Linux"), ("tenancy", "Shared"),
             ("preInstalledSw", "NA"), ("capacitystatus", "Used"), ("operation", "RunInstances"))
    filters = json.dumps([{"Type": "TERM_MATCH", "Field": k, "Value": v} for k, v in terms])
    out = aws("pricing", "get-products", "--service-code", "AmazonEC2", "--filters", filters,
              region="us-east-1", timeout=90) or {}
    prices = []
    for raw in out.get("PriceList") or []:
        item = json.loads(raw) if isinstance(raw, str) else raw
        for term in ((item.get("terms") or {}).get("OnDemand") or {}).values():
            for dim in (term.get("priceDimensions") or {}).values():
                usd = float((dim.get("pricePerUnit") or {}).get("USD") or 0)
                if usd > 0:
                    prices.append(usd)
    if not prices:
        raise Failed(f"no list price for {itype} in {region}")
    return min(prices), "AWS on-demand price, Linux"


# ---------------------------------------------------------------- estimate

def estimate(c: Cluster, pricer: Pricer) -> None:
    """Right-size requests, then count the nodes each pool needs."""
    if c.failed:
        if c.target.skip:
            return
        for p in c.target.pools:
            row = PoolRow(p.name, p.instance_type, p.capacity, p.os, p.arch, p.count)
            row.price, row.note = pricer.hourly(c.cloud, c.region, p.instance_type, p.capacity, p.os)
            c.pools.append(row)
        return
    pods_on = defaultdict(list)
    for p in c.pods:
        pods_on[p.node].append(p)
    groups = defaultdict(list)
    for n in c.nodes:
        groups[(n.pool, n.instance_type, n.capacity)].append(n)
    for (pool, itype, capacity), nodes in sorted(groups.items()):
        count = len(nodes)
        first = nodes[0]
        row = PoolRow(pool, itype, capacity, first.os, first.arch, count,
                      vcpu=first.cpu_cap, mem=first.mem_cap,
                      cpu_alloc=min(n.cpu_alloc for n in nodes), mem_alloc=min(n.mem_alloc for n in nodes),
                      ds_cpu=0.0, ds_mem=0.0, cpu_req=0.0, mem_req=0.0, cpu_rec=0.0, mem_rec=0.0,
                      pods_max=min(n.pods_alloc for n in nodes), pods=0)
        ds_pods = 0
        for n in nodes:
            ds_pods = max(ds_pods, sum(1 for p in pods_on[n.name] if p.kind in OVERHEAD_KINDS))
            for p in pods_on[n.name]:
                cpu, mem = p.request()
                if p.kind in OVERHEAD_KINDS:
                    row.ds_cpu += cpu / count
                    row.ds_mem += mem / count
                else:
                    rcpu, rmem = p.request(recommended=True)
                    row.pods += 1
                    row.cpu_req, row.mem_req = row.cpu_req + cpu, row.mem_req + mem
                    row.cpu_rec, row.mem_rec = row.cpu_rec + rcpu, row.mem_rec + rmem
        net_cpu, net_mem = row.cpu_alloc - row.ds_cpu, row.mem_alloc - row.ds_mem
        net_pods = row.pods_max - ds_pods
        if net_cpu <= 0 or net_mem <= 0 or net_pods <= 0:
            row.needed = count
        else:
            row.needed = min(count, max(math.ceil(row.cpu_rec / (net_cpu * PACKING) - 1e-9),
                                        math.ceil(row.mem_rec / (net_mem * PACKING) - 1e-9),
                                        math.ceil(row.pods / (net_pods * PACKING) - 1e-9)))
        row.price, row.note = pricer.hourly(c.cloud, c.region, itype, capacity, first.os)
        c.pools.append(row)
    keep_minimums(c)


def keep_minimums(c: Cluster) -> None:
    """Keep each pool at its autoscaler minimum, and at least 1 node, across its machine types."""
    minimums = {p.name: p.minimum or 0 for p in c.target.pools}
    no_minimum = {n.pool for n in c.nodes if n.karpenter}
    unknown = []
    for pool in sorted({r.pool for r in c.pools}):
        rows = [r for r in c.pools if r.pool == pool]
        short = max(1, minimums.get(pool, 0)) - sum(r.needed for r in rows)
        while short > 0 and any(r.needed < r.nodes for r in rows):
            max(rows, key=lambda r: r.nodes - r.needed).needed += 1
            short -= 1
        if pool not in minimums and pool not in no_minimum and any(r.needed < r.nodes for r in rows):
            unknown.append(pool)
    if unknown and c.cloud in ("azure", "aws"):
        c.notes.append(f"node pool minimum not known for {', '.join(unknown)}, so nodes needed may be below it")


def money(c: Cluster) -> tuple:
    """(current, optimized, savings) in $/month over priced pools; None when unknown."""
    priced = [r for r in c.pools if r.price is not None]
    if not priced:
        return None, None, None
    current = sum(r.monthly(r.nodes) for r in priced)
    estimated = [r for r in priced if r.needed is not None]
    if not estimated:
        return current, None, None
    savings = sum(max(0.0, r.monthly(r.nodes) - r.monthly(r.needed)) for r in estimated)
    return current, current - savings, savings


# ---------------------------------------------------------------- report rows

class Names:
    """Stable short hashes for names, salted per run."""

    def __init__(self, names: bool, clusters: bool):
        self.names, self.clusters, self.salt = names or clusters, clusters, os.urandom(16)

    def _hash(self, prefix: str, value: str) -> str:
        return prefix + hashlib.sha256(self.salt + value.encode("utf-8")).hexdigest()[:8]

    def cluster(self, v: str) -> str:
        return self._hash("cluster-", v) if self.clusters else v

    def namespace(self, v: str) -> str:
        return self._hash("ns-", v) if self.names else v

    def workload(self, ns: str, v: str) -> str:
        return self._hash("wl-", ns + "/" + v) if self.names else v

    def node(self, v: str) -> str:
        return self._hash("node-", v) if self.names else v


def ratio(a: Optional[float], b: Optional[float]) -> Optional[float]:
    return a / b if a is not None and b else None


def yes_no(v: Optional[bool]) -> Optional[str]:
    return None if v is None else "Yes" if v else "No"


def gib(v: Optional[float]) -> Optional[float]:
    return None if v is None else v / GIB


def usage_label(c: Cluster, minutes: float) -> str:
    if c.failed:
        return ""
    if not c.samples:
        return f"none: {c.metrics_note or 'no metrics'}"
    if c.snapshot:
        return "1 snapshot (command invoke)"
    return f"{c.samples} over {minutes:g} min" if c.samples > 1 else "1 sample"


def cluster_figures(c: Cluster) -> dict:
    alloc_cpu = sum(n.cpu_alloc for n in c.nodes)
    alloc_mem = sum(n.mem_alloc for n in c.nodes)
    req = [p.request() for p in c.pods]
    used = [n for n in c.nodes if n.name in c.node_use] if c.samples else []
    current, optimized, savings = money(c)
    nodes = len(c.nodes) if not c.failed else sum(r.nodes for r in c.pools) or None
    return {
        "nodes": nodes,
        "vcpu": sum(n.cpu_cap for n in c.nodes) if c.nodes else None,
        "mem": sum(n.mem_cap for n in c.nodes) if c.nodes else None,
        "alloc_cpu": alloc_cpu, "alloc_mem": alloc_mem,
        "req_cpu": sum(r[0] for r in req) if c.nodes else None,
        "req_mem": sum(r[1] for r in req) if c.nodes else None,
        "use_cpu": sum(c.node_use[n.name][0] for n in used) if used else None,
        "use_mem": sum(c.node_use[n.name][1] for n in used) if used else None,
        "alloc_cpu_used": sum(n.cpu_alloc for n in used) if used else None,
        "alloc_mem_used": sum(n.mem_alloc for n in used) if used else None,
        "current": current, "optimized": optimized, "savings": savings,
    }


def status(c: Cluster) -> str:
    parts = [c.failed] if c.failed else []
    if c.failed and c.pools:
        parts.append("nodes and cost from the cloud API only")
    parts += c.notes
    unpriced = [r for r in c.pools if r.price is None]
    if unpriced:
        notes = ", ".join(sorted({r.note for r in unpriced}))
        parts.append(f"{sum(r.nodes for r in unpriced)} node(s) not priced: {notes}")
    return "; ".join(parts) or "OK"


SUMMARY_COLS = [("Cluster", "text"), ("Cloud", "text"), ("Region", "text"), ("K8s version", "text"),
                ("Nodes", "int"), ("vCPU", "dec"), ("Memory GiB", "dec"), ("CPU requested %", "pct"),
                ("CPU used %", "pct"), ("Memory requested %", "pct"), ("Memory used %", "pct"),
                ("Current monthly $", "usd"), ("Estimated optimized monthly $", "usd"),
                ("Estimated monthly savings $", "usd"), ("Savings %", "pct"), ("Usage samples", "text"),
                ("Status", "text")]
POOL_COLS = [("Cluster", "text"), ("Node pool", "text"), ("Instance type", "text"), ("Capacity", "text"),
             ("OS", "text"), ("Arch", "text"), ("Nodes", "int"), ("vCPU per node", "dec"),
             ("Memory GiB per node", "dec"), ("Allocatable vCPU per node", "dec"),
             ("Allocatable memory GiB per node", "dec"), ("DaemonSet vCPU per node", "dec"),
             ("DaemonSet memory GiB per node", "dec"), ("Workload vCPU requested", "dec"),
             ("Workload vCPU recommended", "dec"), ("Workload memory GiB requested", "dec"),
             ("Workload memory GiB recommended", "dec"), ("Max pods per node", "int"), ("Workload pods", "int"),
             ("Nodes needed", "int"), ("Price $/hour per node", "usd4"),
             ("Current monthly $", "usd"), ("Estimated optimized monthly $", "usd"),
             ("Estimated monthly savings $", "usd"), ("Price note", "text")]
NODE_COLS = [("Cluster", "text"), ("Node", "text"), ("Node pool", "text"), ("Instance type", "text"),
             ("Capacity", "text"), ("Zone", "text"), ("Ready", "text"), ("Pods", "int"),
             ("Allocatable vCPU", "dec"), ("vCPU requested", "dec"), ("vCPU used (peak)", "dec"),
             ("CPU requested %", "pct"), ("CPU used %", "pct"), ("Allocatable memory GiB", "dec"),
             ("Memory requested GiB", "dec"), ("Memory used GiB (peak)", "dec"), ("Memory requested %", "pct"),
             ("Memory used %", "pct")]
WORKLOAD_COLS = [("Cluster", "text"), ("Namespace", "text"), ("Kind", "text"), ("Name", "text"),
                 ("Node pool", "text"), ("Replicas", "int"), ("CPU request m", "int"), ("CPU limit m", "int"),
                 ("CPU used m (peak)", "int"), ("CPU recommended m", "int"), ("Memory request MiB", "int"),
                 ("Memory limit MiB", "int"), ("Memory used MiB (peak)", "int"), ("Memory recommended MiB", "int"),
                 ("HPA", "text"), ("PDB", "text"), ("RWO volume", "text")]


def summary_rows(clusters: list, names: Names, minutes: float) -> tuple:
    rows, tot = [], defaultdict(float)
    seen = defaultdict(bool)
    for c in clusters:
        f = cluster_figures(c)
        for k, v in f.items():
            if v is not None:
                tot[k] += v
                seen[k] = True
        rows.append([names.cluster(c.target.name), {"aws": "AWS", "azure": "Azure"}.get(c.cloud, c.cloud.capitalize()),
                     c.region, re.sub(r"^v", "", c.version), f["nodes"], f["vcpu"], gib(f["mem"]),
                     ratio(f["req_cpu"], f["alloc_cpu"]), ratio(f["use_cpu"], f["alloc_cpu_used"]),
                     ratio(f["req_mem"], f["alloc_mem"]), ratio(f["use_mem"], f["alloc_mem_used"]),
                     f["current"], f["optimized"], f["savings"], ratio(f["savings"], f["current"]),
                     usage_label(c, minutes), status(c)])

    def t(key):
        return tot[key] if seen[key] else None

    current, savings = t("current"), t("savings")
    optimized = None if current is None or savings is None else current - savings
    total = ["TOTAL", "", "", "", t("nodes"), t("vcpu"), gib(t("mem")),
             ratio(t("req_cpu"), tot["alloc_cpu"]), ratio(t("use_cpu"), t("alloc_cpu_used")),
             ratio(t("req_mem"), tot["alloc_mem"]), ratio(t("use_mem"), t("alloc_mem_used")),
             current, optimized, savings, ratio(savings, current), "", ""]
    return rows, total


def pool_rows(clusters: list, names: Names) -> list:
    rows = []
    for c in clusters:
        for r in c.pools:
            optimized = r.monthly(r.needed)
            current = r.monthly(r.nodes)
            rows.append([names.cluster(c.target.name), r.pool, r.instance_type, r.capacity.capitalize(), r.os, r.arch,
                         r.nodes, r.vcpu, gib(r.mem), r.cpu_alloc, gib(r.mem_alloc), r.ds_cpu, gib(r.ds_mem),
                         r.cpu_req, r.cpu_rec, gib(r.mem_req), gib(r.mem_rec), r.pods_max, r.pods, r.needed, r.price,
                         current, optimized, None if optimized is None else max(0.0, current - optimized),
                         ("count from the cloud API; " if c.failed else "") + r.note])
    return rows


def node_rows(clusters: list, names: Names) -> list:
    rows = []
    for c in clusters:
        on = defaultdict(list)
        for p in c.pods:
            on[p.node].append(p.request())
        for n in sorted(c.nodes, key=lambda n: (n.pool, n.name)):
            use = c.node_use.get(n.name) if c.samples else None
            cpu_req, mem_req = sum(r[0] for r in on[n.name]), sum(r[1] for r in on[n.name])
            rows.append([names.cluster(c.target.name), names.node(n.name), n.pool, n.instance_type,
                         n.capacity.capitalize(), n.zone, yes_no(n.ready), len(on[n.name]),
                         n.cpu_alloc, cpu_req, use[0] if use else None, ratio(cpu_req, n.cpu_alloc),
                         ratio(use[0], n.cpu_alloc) if use else None, gib(n.mem_alloc), gib(mem_req),
                         gib(use[1]) if use else None, ratio(mem_req, n.mem_alloc),
                         ratio(use[1], n.mem_alloc) if use else None])
    return rows


def workload_rows(clusters: list, names: Names) -> list:
    rows = []
    for c in clusters:
        pool_of = {n.name: n.pool for n in c.nodes}
        for w in c.workloads:
            req = [p.request() for p in w.pods]
            rec = [p.request(recommended=True) for p in w.pods]
            lims = [p.limits() for p in w.pods]
            cpu_lim = None if any(x[0] is None for x in lims) else sum(x[0] for x in lims)
            mem_lim = None if any(x[1] is None for x in lims) else sum(x[1] for x in lims)
            ctrs = [x for p in w.pods for x in p.running()]
            complete = bool(ctrs) and all(x.cpu_use is not None for x in ctrs)
            cpu_use = sum(x.cpu_use for x in ctrs) if complete else None
            mem_use = sum(x.mem_use for x in ctrs) if complete else None
            rows.append([names.cluster(c.target.name), names.namespace(w.namespace), w.kind,
                         names.workload(w.namespace, w.name),
                         ", ".join(sorted({pool_of.get(p.node, "") for p in w.pods})), len(w.pods),
                         milli(sum(r[0] for r in req)), milli(cpu_lim), milli(cpu_use), milli(sum(r[0] for r in rec)),
                         mib(sum(r[1] for r in req)), mib(mem_lim), mib(mem_use), mib(sum(r[1] for r in rec)),
                         yes_no(w.hpa), yes_no(w.pdb), yes_no(w.rwo)])
    return rows


def milli(v: Optional[float]) -> Optional[int]:
    return None if v is None else int(round(v * 1000))


def mib(v: Optional[float]) -> Optional[int]:
    return None if v is None else int(round(v / MIB))


def assumptions(minutes: float, anonymized: bool) -> list:
    rows = [
        ("What this is", "An estimate, not a quote. Zoorik prepared this method; the figures come from your clusters "
                         "and public list prices at the time of the run."),
        ("Step 1: right-size", f"Recommended request per container = peak sampled usage x {HEADROOM:g}, at least "
                               f"{CPU_FLOOR * 1000:g}m CPU and {MEM_FLOOR // MIB} MiB memory. It never goes above "
                               "today's request, unless today's request is zero. With no usage data, today's request "
                               "is kept. Init containers keep their requests, except native sidecars (restartPolicy: "
                               "Always), which are right-sized like the other running containers. A pod with "
                               "pod-level requests keeps them as they are."),
        ("Step 2: remove nodes", "Per node pool and machine type: DaemonSet and static pods are per-node overhead "
                                 "(today's requests). Nodes needed = the larger of ceil(recommended CPU / "
                                 f"(allocatable CPU - overhead) / {PACKING:g}) and the same for memory, and never "
                                 "above today's count. A node also holds at most its allocatable pod count, less "
                                 "the DaemonSet and static pods on it, so nodes needed is never below ceil(workload "
                                 f"pods / that number / {PACKING:g}). Each pool keeps at least its autoscaler "
                                 "minimum, and at least 1 node, across its machine types. Where the cloud API gives "
                                 "no minimum (a cluster read through a kube context, or a node group the API does "
                                 "not list), the floor is 1 node and the Status column says so."),
        ("Savings", "Estimated monthly savings = (nodes today - nodes needed) x node price x "
                    f"{HOURS_PER_MONTH} hours. Clusters or pools with no estimate add their cost and no savings "
                    "to the totals."),
        ("Usage", "Peak of kubectl top (metrics-server), "
                  + (f"sampled every {SAMPLE_INTERVAL} s for {minutes:g} min. " if minutes else "one sample. ")
                  + "Clusters read through az aks command invoke have one snapshot. A short window can miss peaks; "
                  "Zoorik's own recommendations use a longer history."),
        ("Prices", "Azure: the public Azure Retail Prices API, Linux pay-as-you-go, or the Spot meter for Spot "
                   "nodes. AWS: the AWS Price List API, Linux, shared tenancy, on-demand; Spot nodes use the current "
                   "Spot price averaged across zones. USD list prices, without discounts, reservations, savings "
                   f"plans or credits. A month is {HOURS_PER_MONTH} hours. Nodes whose labels don't say Spot or "
                   "on-demand (such as self-managed EKS node groups) are not priced."),
        ("Not included", "Moving to Spot, cheaper or newer machine types, and packing across pools. These are "
                         "further savings on top of this estimate. Also not counted: storage, network, load "
                         "balancers, control-plane fees, licences (Windows nodes are not priced), and pending pods."),
        ("Blank cells", "A blank figure was not available; the Status or Price note column says why. A blank limit "
                        "means at least one container has no limit."),
        ("Data read", "Nodes, pods, ReplicaSets, Jobs, HPAs, PDBs and volume claims (kubectl get), and usage "
                      "(kubectl top). Kept: names, owners, requests, limits, labels for matching and volume claim "
                      "names. Never read: Secrets, ConfigMaps. Never kept: env values, annotations, pod IPs."),
        ("What ran", "kubectl get and kubectl top in each cluster; the az and aws CLIs to read the signed-in "
                     "account, list clusters, node pools and regions, write each cluster's credentials to a "
                     "temporary kubeconfig and look up AWS prices; and the public Azure price list. For a private "
                     "AKS cluster kubectl could not read from where the script ran, or an AKS cluster kubectl could "
                     "not reach or sign in to, Azure ran the same kubectl reads in a short-lived pod of its own, in "
                     "the cluster's aks-command namespace (az aks command invoke, recorded in Azure's activity log; "
                     "--no-invoke skips this). The script sent nothing to Zoorik."),
    ]
    if anonymized:
        rows.append(("Anonymized", "Namespace, workload and node names are replaced by short hashes salted for this "
                                   "run only, so they cannot be reversed."))
    rows.append(("Prepared by", f"Zoorik Kubernetes assessment {VERSION}. Run it again: {ONE_LINER}  "
                                f"Questions: {SEND_TO}"))
    return [list(r) for r in rows]


# ---------------------------------------------------------------- xlsx writer

_FORMATS = {"text": 0, "int": 3, "dec": 2, "pct": 164, "usd": 165, "usd4": 166}


def _cell_styles() -> tuple:
    """(numFmt, font, fill, border, wrap) per style, and the index of each style."""
    xfs = [(0, 0, 0, 0, False), (0, 1, 2, 1, False), (0, 2, 0, 0, False), (0, 0, 0, 0, True)]
    index = {"header": 1, "title": 2, "wrap": 3}
    for bold in (False, True):
        for fmt, num in _FORMATS.items():
            index[(fmt, bold)] = len(xfs)
            xfs.append((num, 1 if bold else 0, 0, 0, False))
    return xfs, index


_XFS, STYLE = _cell_styles()

_NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
_RNS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
_BAD_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")


@dataclass
class Sheet:
    name: str
    columns: list
    rows: list
    preamble: list = field(default_factory=list)
    total: Optional[list] = None
    widths: Optional[list] = None
    autofilter: bool = True


def col_letter(i: int) -> str:
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def cell_xml(ref: str, value, style: int) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if value != value or value in (float("inf"), float("-inf")):
            return ""
        return f'<c r="{ref}" s="{style}"><v>{value!r}</v></c>'
    text = escape(_BAD_XML.sub("", str(value)))
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def shown_length(value, fmt: str) -> int:
    if value is None:
        return 0
    if fmt == "int":
        return len(f"{value:,.0f}")
    if fmt == "dec":
        return len(f"{value:,.2f}")
    if fmt == "pct":
        return len(f"{value * 100:.1f}%")
    if fmt == "usd":
        return len(f"${value:,.0f}")
    if fmt == "usd4":
        return len(f"${value:.4f}")
    return len(str(value))


def sheet_xml(s: Sheet) -> tuple:
    head = len(s.preamble) + 1
    last_col = col_letter(len(s.columns) - 1)
    rows = s.rows + ([s.total] if s.total else [])
    widths = s.widths or [min(60, max([len(name)] + [shown_length(r[i], fmt) for r in rows]) + 2)
                          for i, (name, fmt) in enumerate(s.columns)]
    out = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet {_NS} {_RNS}>',
           '<sheetViews><sheetView workbookViewId="0" showGridLines="0">'
           f'<pane ySplit="{head}" topLeftCell="A{head + 1}" activePane="bottomLeft" state="frozen"/>'
           '<selection pane="bottomLeft"/></sheetView></sheetViews>',
           '<sheetFormatPr defaultRowHeight="15"/><cols>']
    out += [f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths)]
    out.append("</cols><sheetData>")
    for r, (text, style) in enumerate(s.preamble, 1):
        out.append(f'<row r="{r}">{cell_xml(f"A{r}", text, STYLE[style] if style else 0)}</row>')
    out.append(f'<row r="{head}">' + "".join(cell_xml(f"{col_letter(i)}{head}", name, STYLE["header"])
                                             for i, (name, _) in enumerate(s.columns)) + "</row>")
    body = [(row, False) for row in s.rows] + ([(s.total, True)] if s.total else [])
    for n, (row, bold) in enumerate(body, head + 1):
        cells = []
        for i, (value, (_, fmt)) in enumerate(zip(row, s.columns)):
            style = STYLE["wrap"] if s.widths and fmt == "text" and i == len(s.columns) - 1 else STYLE[(fmt, bold)]
            cells.append(cell_xml(f"{col_letter(i)}{n}", value, style))
        out.append(f'<row r="{n}">{"".join(cells)}</row>')
    out.append("</sheetData>")
    ref = f"A{head}:{last_col}{head + len(s.rows)}" if s.autofilter and s.rows else ""
    if ref:
        out.append(f'<autoFilter ref="{ref}"/>')
    out.append("</worksheet>")
    return "".join(out), ref


def styles_xml() -> str:
    xfs = "".join(
        f'<xf numFmtId="{num}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0" applyNumberFormat="1" '
        'applyFont="1" applyFill="1" applyBorder="1"'
        + ('><alignment wrapText="1" vertical="top"/></xf>' if wrap else "/>")
        for num, font, fill, border, wrap in _XFS)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<styleSheet {_NS}><numFmts count="3"><numFmt numFmtId="164" formatCode="0.0%"/>'
            '<numFmt numFmtId="165" formatCode="$#,##0"/><numFmt numFmtId="166" formatCode="$0.0000"/></numFmts>'
            '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/>'
            '<name val="Calibri"/></font><font><b/><sz val="14"/><name val="Calibri"/></font></fonts>'
            '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/>'
            '</fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8EEF7"/><bgColor indexed="64"/>'
            '</patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
            '<border><left/><right/><top/><bottom style="thin"><color rgb="FF9AA5B1"/></bottom><diagonal/></border>'
            '</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            f'<cellXfs count="{len(_XFS)}">{xfs}</cellXfs><cellStyles count="1">'
            '<cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')


def write_xlsx(path: str, sheets: list) -> None:
    """Write the sheets as a minimal .xlsx with the standard library only."""
    sheets = [s for s in sheets if s.rows]
    parts, names = {}, []
    for i, s in enumerate(sheets, 1):
        xml, ref = sheet_xml(s)
        parts[f"xl/worksheets/sheet{i}.xml"] = xml
        if ref:
            absolute = re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", ref)
            names.append(f'<definedName name="_xlnm._FilterDatabase" localSheetId="{i - 1}" hidden="1">'
                         f"'{s.name}'!{absolute}</definedName>")
    sheet_list = "".join(f'<sheet name="{escape(s.name)}" sheetId="{i}" r:id="rId{i}"/>'
                         for i, s in enumerate(sheets, 1))
    parts["xl/workbook.xml"] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook {_NS} {_RNS}>'
                                f'<bookViews><workbookView/></bookViews><sheets>{sheet_list}</sheets>'
                                + (f"<definedNames>{''.join(names)}</definedNames>" if names else "") + "</workbook>")
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    parts["xl/_rels/workbook.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(f'<Relationship Id="rId{i}" Type="{rel}/worksheet" Target="worksheets/sheet{i}.xml"/>'
                  for i in range(1, len(sheets) + 1))
        + f'<Relationship Id="rId{len(sheets) + 1}" Type="{rel}/styles" Target="styles.xml"/></Relationships>')
    parts["xl/styles.xml"] = styles_xml()
    parts["_rels/.rels"] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                            f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/>'
                            '</Relationships>')
    ct = "application/vnd.openxmlformats-officedocument.spreadsheetml"
    parts["[Content_Types].xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'<Override PartName="/xl/workbook.xml" ContentType="{ct}.sheet.main+xml"/>'
        f'<Override PartName="/xl/styles.xml" ContentType="{ct}.styles+xml"/>'
        + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="{ct}.worksheet+xml"/>'
                  for i in range(1, len(sheets) + 1)) + "</Types>")
    order = ["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/styles.xml"]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in order + sorted(k for k in parts if k not in order):
            z.writestr(name, parts[name])


# ---------------------------------------------------------------- main

def parse_args(argv: Optional[list]) -> argparse.Namespace:
    p = argparse.ArgumentParser(prog="coral.py", usage=f"{ONE_LINER} [options]",
                                description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--context", action="append", metavar="NAME", help="assess this kube context (repeatable)")
    p.add_argument("--all-contexts", action="store_true", help="assess every context in your kubeconfig")
    p.add_argument("--discover", action="append", choices=("azure", "aws"),
                   help="find clusters with az (AKS) or aws (EKS); repeatable")
    p.add_argument("--subscription", metavar="ID", help="Azure subscription to search (implies --discover azure)")
    p.add_argument("--region", action="append", metavar="NAME",
                   help="AWS region to search, repeatable (implies --discover aws; default: every enabled region)")
    p.add_argument("--sample-minutes", type=float, default=3, metavar="N",
                   help="minutes to sample usage (default 3; 0 takes one sample)")
    p.add_argument("--no-prices", action="store_true", help="skip price lookups")
    p.add_argument("--no-invoke", action="store_true",
                   help="never use az aks command invoke (AKS clusters kubectl can't read are then not assessed)")
    p.add_argument("--anonymize", action="store_true", help="replace namespace, workload and node names with hashes")
    p.add_argument("--anonymize-clusters", action="store_true", help="also replace cluster names (implies --anonymize)")
    p.add_argument("--out", metavar="PATH", help="output file or folder")
    p.add_argument("--version", action="version", version=f"Zoorik Kubernetes assessment {VERSION}")
    args = p.parse_args(argv)
    if not math.isfinite(args.sample_minutes) or args.sample_minutes < 0:
        p.error("--sample-minutes must be a number, 0 or more")
    return args


def output_path(out: Optional[str], cloud_shell: str) -> str:
    """Choose the report's path before any work, so a folder that can't be written is found early."""
    name = f"zoorik-k8s-assessment-{datetime.now(timezone.utc):%Y%m%d-%H%M}.xlsx"
    if out:
        if out.endswith(("/", os.sep)) or os.path.isdir(out):
            path = os.path.join(out, name)
        else:
            path = out if out.lower().endswith(".xlsx") else out + ".xlsx"
        folder = os.path.dirname(os.path.abspath(path))
        try:
            os.makedirs(folder, exist_ok=True)
        except OSError as e:
            raise Failed(f"cannot create {folder}: {e.strerror}") from None
        if not os.access(folder, os.W_OK):
            raise Failed(f"cannot write to {folder}")
        return os.path.abspath(path)
    home = os.path.expanduser("~")
    for folder in [home] if cloud_shell else [os.getcwd(), home]:
        if os.access(folder, os.W_OK):
            return os.path.join(folder, name)
    raise Failed("cannot write to the current folder or the home folder; choose one with --out PATH")


def unexpected(e: Exception) -> str:
    """Name the error and the last line of this script it passed through, for the report to Zoorik."""
    here = unexpected.__code__.co_filename
    line = next((f.lineno for f in reversed(traceback.extract_tb(e.__traceback__)) if f.filename == here), 0)
    return f"unexpected error ({type(e).__name__}, line {line})"


def assess(args: argparse.Namespace) -> int:
    say(f"Zoorik Kubernetes assessment {VERSION}. It sends nothing to Zoorik; you send the report yourself.")
    shell = "azure" if in_azure_cloud_shell() else "aws" if in_aws_cloudshell() else ""
    try:
        path = output_path(args.out, shell)
    except Failed as e:
        say(f"Cannot write the report: {e}")
        return 1
    targets, search_failed, why = choose_targets(args, shell)
    if not targets:
        if shell == "azure":
            say(("No cluster to assess." if search_failed else "No AKS cluster found in this subscription.")
                + f" To search another subscription: {ONE_LINER} --subscription ID")
        elif shell == "aws":
            say("No cluster to assess." if search_failed else "No EKS cluster found in this AWS account.")
        else:
            say(f"No cluster found{f' ({why})' if why else ''}: run this in Azure Cloud Shell or AWS CloudShell, "
                "or point kubectl at a cluster (kubectl config use-context NAME).")
        return 1
    if shell and (len(targets) > 5 or sum(t.source == "aks" and t.private for t in targets) > 1
                  or args.sample_minutes > 10):
        say("This run can take a while. Cloud Shell closes a session after 20 minutes without a key press, "
            "so press Enter now and then until the report path appears.")
    clusters, seen = [], []
    for i, t in enumerate(targets, 1):
        say(f"[{i}/{len(targets)}] {t.name}{f' ({t.cloud}, {t.region})' if t.cloud else ''}: ", end="")
        try:
            c = collect(t, not args.no_invoke)
        except Exception as e:
            c = Cluster(target=t, cloud=t.cloud, region=t.region, version=t.version, failed=unexpected(e))
        ids = {n.provider for n in c.nodes if n.provider}
        same = next((name for name, other in seen if ids & other), None)
        if same:
            say(f"same cluster as {same}; counted once")
            continue
        seen.append((t.name, ids))
        clusters.append(c)
        say(f"not assessed, {c.failed}" if c.failed else f"{len(c.nodes)} nodes, {len(c.pods)} pods")
    sample_all(clusters, args.sample_minutes)
    if not args.no_prices:
        say("Looking up list prices ...")
    pricer = Pricer(not args.no_prices)
    for c in clusters:
        apply_usage(c)
        estimate(c, pricer)

    names = Names(args.anonymize, args.anonymize_clusters)
    rows, total = summary_rows(clusters, names, args.sample_minutes)
    generated = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    sheets = [
        Sheet("Summary", SUMMARY_COLS, rows, total=total,
              preamble=[("Zoorik Kubernetes cost assessment", "title"), (f"Generated {generated}", ""),
                        (f"Script version {VERSION}", ""), ("", "")]),
        Sheet("Node pools", POOL_COLS, pool_rows(clusters, names)),
        Sheet("Nodes", NODE_COLS, node_rows(clusters, names)),
        Sheet("Workloads", WORKLOAD_COLS, workload_rows(clusters, names)),
        Sheet("Assumptions", [("Topic", "text"), ("Detail", "text")],
              assumptions(args.sample_minutes, names.names), widths=[22, 110], autofilter=False),
    ]
    try:
        write_xlsx(path, sheets)
    except OSError as e:
        say(f"Could not write {path}: {e.strerror}")
        return 1

    assessed = sum(1 for c in clusters if not c.failed)
    missed = len(clusters) - assessed
    current, savings = total[11], total[13]
    nodes = int(total[4] or 0)
    priced = sum(r.nodes for c in clusters for r in c.pools if r.price is not None)
    cost = f"${current:,.0f} at list prices" if current is not None else "not available"
    if current is not None and priced < nodes:
        cost += f", for {priced} of {nodes} nodes (the rest are not priced; see Status)"
    saved = f"${savings:,.0f} per month ({savings / current:.0%})" if current and savings is not None else "not available"
    say()
    say(f"Clusters:          {assessed} assessed" + (f", {missed} not assessed" if missed else ""))
    say(f"Monthly cost:      {cost}")
    say(f"Estimated savings: {saved}")
    say()
    say(f"Report: {path}")
    if shell == "azure":
        say(f"Download it: Cloud Shell toolbar > Manage files > Download, then enter {path}")
    elif shell == "aws":
        say(f"Download it: Actions > Download file, then enter {path}")
        say("(Download needs the full CloudShell page; if Actions is missing, open CloudShell in a new tab. "
            "A CloudShell VPC environment can't download files.)")
    say(f"Please e-mail it to {SEND_TO}")
    return 0


def stop(signum, _frame) -> None:
    """Turn a hang-up or terminate into a normal exit, so the temporary kubeconfig is removed."""
    raise SystemExit(128 + signum)


def main(argv: Optional[list] = None) -> int:
    global WORKDIR
    args = parse_args(argv)
    for name in ("SIGHUP", "SIGTERM"):
        if hasattr(signal, name):
            signal.signal(getattr(signal, name), stop)
    WORKDIR = tempfile.mkdtemp(prefix="zoorik-k8s-")
    try:
        return assess(args)
    except KeyboardInterrupt:
        say("\nStopped.")
        return 130
    except Exception as e:
        say(f"\nStopped by an {unexpected(e)}. Please e-mail this line to {SEND_TO}.")
        return 1
    finally:
        shutil.rmtree(WORKDIR, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
