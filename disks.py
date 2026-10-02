#!/usr/bin/env python3
"""Zoorik ElasticVol disk assessment.

Reads the Azure VMs and managed disks you can see, estimates what Zoorik
ElasticVol (Zoorik block storage) could save by right-sizing the data disks of
your Linux VMs, and writes one Excel file for you to e-mail to Zoorik.
Nothing to install.

Run it in Azure Cloud Shell, or in any shell where az is signed in:
    curl -fsSL https://zoorik.com/assessments/disks | python3 -

The source is at github.com/Zoorikcloud/assessment.

What it does. Apart from the one step marked below, which runs only if you
answer y, it only reads, and it sends nothing anywhere.
  - Azure Resource Graph: your VMs, managed disks and resource groups, in
    every enabled subscription of the directory you are signed in to.
  - Azure Compute: each VM size's data-disk slots and disk limits, and where
    Premium SSD v2 is offered.
  - Azure Monitor: 14 days of IOPS and throughput for each data disk of a
    Linux VM.
  - prices.azure.com: the public list prices of managed disks.
  - Only if you answer y to its one question: one command that only reads,
    inside each running Linux VM with data disks, through Azure Run Command.
    It reads used space (df, lsblk), the kernel version, the Linux
    distribution and whether btrfs is available. On its side, Azure adds its
    Run Command extension to a VM that does not have it yet, keeps a copy of
    the command and its output under /var/lib/waagent on the VM, and records
    each run in the activity log.
It writes one file: the Excel report, in the current directory.
"""

import json
import math
import os
import re
import subprocess
import sys

if sys.version_info < (3, 8):
    print("This needs Python 3.8 or later. Open Azure Cloud Shell (shell.azure.com) and paste the line there.")
    sys.exit(1)

import http.client
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import Callable, Dict, List, Optional, Tuple
from xml.sax.saxutils import escape

# ---------------------------------------------------------------- config

VERSION = "1.0.0"
ONE_LINER = "curl -fsSL https://zoorik.com/assessments/disks | python3 -"
SEND_TO = "founders@zoorik.com"

HOURS_PER_MONTH = 730
GIB = 2 ** 30
METRIC_DAYS = 14
WORKERS = 8
RUN_COMMAND_TIMEOUT = 300
RUN_COMMAND_SECONDS = 45
RUN_COMMAND_PERMISSION = "Microsoft.Compute/virtualMachines/runCommand/action"
PRICES_URL = "https://prices.azure.com/api/retail/prices"
GRAPH_URL = "/providers/Microsoft.ResourceGraph/resources?api-version=2022-10-01"
SKUS_API = "2021-07-01"
METRICS_API = "2023-10-01"
METRICS = ("Composite Disk Read Operations/sec", "Composite Disk Write Operations/sec",
           "Composite Disk Read Bytes/sec", "Composite Disk Write Bytes/sec")

# How Zoorik ElasticVol sizes a pool: the same rules and numbers the product plans with.
HEADROOM_FRACTION = 0.20
HEADROOM_MIN_GIB = 50
ANCHOR_POOL_THRESHOLD_GIB = 512
ANCHOR_COUNT = 2
ANCHOR_CAP_FRACTION = 0.75
ANCHOR_MAX_GIB = 2048
MIN_ELASTICS = 3
MAX_ELASTICS = 6
ELASTIC_FLOOR_V2_GIB = 64
ELASTIC_FLOOR_V1_GIB = 128
ELASTIC_CAP_GIB = 512
ELASTIC_DIVISOR = 16
ELASTIC_FILL_FRACTION = 0.80
EVAC_SHARE = 0.30
EVAC_WINDOW_MINUTES = 120
SMALL_MEMBER_GIB = 8
DEGRADED_ANCHORS = 1
DEGRADED_MAX_ELASTICS = 3
MIN_POOL_SLOTS = 2
MIN_VM_SLOTS = 8
RESERVED_SLOTS = 1
DEMAND_FACTOR = 1.3
KERNEL_FLOOR = (5, 14)

# Premium SSD v2 limits, from Azure's documentation.
V2_MIN_IOPS = 3000
V2_MIN_MBPS = 125
V2_MAX_IOPS_PER_GIB = 500
V2_MBPS_PER_IOPS = 0.25
V2_MAX_MBPS = 2000

# The pool goes on Premium SSD v2 wherever the VM's region and zone offer it, otherwise on Premium SSD.
POOL_V2 = "Premium SSD v2"
POOL_V1 = "Premium SSD"
# Regions where Azure takes Premium SSD v2 on a VM without an availability zone (nonzonal disks), from
# Microsoft Learn, "Deploy a Premium SSD v2 managed disk", Regional availability (updated 10 Sept 2026).
# Azure Compute lists only the zones that offer Premium SSD v2, not this, so a region missing here gets Premium SSD.
V2_NONZONAL_REGIONS = frozenset((
    "australiacentral2", "australiasoutheast", "brazilsoutheast", "canadaeast", "germanynorth", "northcentralus",
    "norwaywest", "southindia", "switzerlandwest", "taiwannorth", "ukwest", "usgovarizona", "westcentralus",
    "westus", "australiaeast", "brazilsouth", "eastasia", "francecentral", "germanywestcentral", "italynorth",
    "koreacentral", "mexicocentral", "newzealandnorth", "norwayeast", "polandcentral", "spaincentral",
    "southafricanorth", "southeastasia", "swedencentral", "switzerlandnorth"))

TIER_SIZES = ((4, "1"), (8, "2"), (16, "3"), (32, "4"), (64, "6"), (128, "10"), (256, "15"), (512, "20"),
              (1024, "30"), (2048, "40"), (4096, "50"), (8192, "60"), (16384, "70"), (32767, "80"))
HDD_TIER_SIZES = TIER_SIZES[3:]
V1_TIER_GIB = tuple(size for size, _ in TIER_SIZES)
# Provisioned IOPS and MB/s of each Premium SSD size, from Azure's documentation (bursting not counted).
V1_TIER_PERFORMANCE = {4: (120, 25), 8: (120, 25), 16: (120, 25), 32: (120, 25), 64: (240, 50), 128: (500, 100),
                       256: (1100, 125), 512: (2300, 150), 1024: (5000, 200), 2048: (7500, 250),
                       4096: (7500, 250), 8192: (16000, 500), 16384: (18000, 750), 32767: (20000, 900)}
DISK_TYPES = {
    "premium_lrs": ("P", "LRS", "Premium SSD"),
    "premium_zrs": ("P", "ZRS", "Premium SSD ZRS"),
    "standardssd_lrs": ("E", "LRS", "Standard SSD"),
    "standardssd_zrs": ("E", "ZRS", "Standard SSD ZRS"),
    "standard_lrs": ("S", "LRS", "Standard HDD"),
    "premiumv2_lrs": ("V2", "LRS", "Premium SSD v2"),
    "ultrassd_lrs": ("ULTRA", "LRS", "Ultra Disk"),
}
TIER_PRODUCTS = ("Premium SSD Managed Disks", "Standard SSD Managed Disks", "Standard HDD Managed Disks")
PROVISIONED_PRODUCTS = {"Azure Premium SSD v2": "Premium LRS ", "Ultra Disks": "Ultra LRS "}
PROVISIONED_METERS = {"Provisioned Capacity": "capacity", "Provisioned IOPS": "iops",
                      "Provisioned Throughput (MBps)": "mbps"}
RHEL_FAMILY = ("rhel", "centos", "rocky", "almalinux", "ol")
OS_NAMES = {"ubuntu": "Ubuntu", "debian": "Debian", "rhel": "RHEL", "centos": "CentOS", "rocky": "Rocky Linux",
            "almalinux": "AlmaLinux", "ol": "Oracle Linux", "sles": "SLES", "sles_sap": "SLES for SAP",
            "opensuse-leap": "openSUSE Leap", "azurelinux": "Azure Linux", "mariner": "CBL-Mariner",
            "fedora": "Fedora", "flatcar": "Flatcar"}

# The one command that runs inside a VM, and only after a yes. It only reads.
GUEST_SCRIPT = r"""echo S
echo "K $(uname -r)"
grep -E '^(ID|ID_LIKE|VERSION_ID)=' /etc/os-release
if grep -qw btrfs /proc/filesystems; then echo "B loaded"; elif modinfo -n btrfs >/dev/null 2>&1; then echo "B module"; else echo "B absent"; fi
for l in /dev/disk/azure/scsi1/lun*; do n=${l##*/lun}; case $n in ''|*[!0-9]*) continue;; esac; echo "L $n $(lsblk -nro MOUNTPOINT "$l" | tr '\n' ' ')"; done
timeout 30 df -P -B1 -T -l | awk '$1 ~ /^\/dev\// && $1 !~ /^\/dev\/loop/ {m=$7; for (i=8; i<=NF; i++) m=m" "$i; print "D", $2, $3, $4, m}'
echo E"""


def say(text: str = "", end: str = "\n") -> None:
    print(text, end=end, flush=True)


# ---------------------------------------------------------------- commands

class Failed(Exception):
    """An expected failure, carried as a short reason."""


_REASONS = (
    ("runcommand/action", "no permission to use Run Command (" + RUN_COMMAND_PERMISSION + ")"),
    ("please run 'az login'", "not signed in to Azure; run az login"),
    ("az login", "Azure sign-in has expired; run az login"),
    ("failed to connect to msi", "Cloud Shell sign-in issue; run az login"),
    ("authorizationfailed", "access denied"),
    ("forbidden", "access denied"),
    ("execution is in progress", "another Run Command is already running on this VM"),
    ("vmagentstatuscommunicationerror", "the VM agent is not reporting"),
    ("vm agent", "the VM agent is not ready"),
    ("deallocat", "the VM is not running"),
    ("vmextensionprovisioning", "Azure could not start its Run Command extension on the VM"),
    ("vmextensionhandler", "Azure could not start its Run Command extension on the VM"),
    ("toomanyrequests", "Azure is throttling requests; try again later"),
    ("subscriptionnotfound", "subscription not found"),
    ("resourcenotfound", "not found"),
    ("timed out", "no answer in time"),
)


def reason(text: str) -> str:
    """Turn tool error output into a short reason, free of user and object names."""
    low = (text or "").lower()
    for needle, short in _REASONS:
        if needle in low:
            return short
    line = next((x.strip() for x in (text or "").splitlines() if x.strip()), "unknown error")
    line = re.sub(r"^(error|ERROR|Error)[:\s]+", "", line)
    line = re.sub(r"\"[^\"]*\"|'[^']*'", "...", line)
    return line[:100]


def az(args: List[str], timeout: float = 120):
    """Run one az command and return its JSON output, or raise Failed. Every Azure call goes through here.

    The Azure CLI's own usage telemetry is switched off for these calls only; your az settings stay as they are."""
    for attempt in range(3):
        try:
            p = subprocess.run(["az"] + list(args) + ["-o", "json", "--only-show-errors"], stdin=subprocess.DEVNULL,
                               stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout,
                               encoding="utf-8", errors="replace",
                               env=dict(os.environ, AZURE_CORE_COLLECT_TELEMETRY="false"))
        except FileNotFoundError:
            raise Failed("the Azure CLI (az) is not installed") from None
        except subprocess.TimeoutExpired:
            raise Failed(f"no answer within {max(1, round(timeout / 60))} min") from None
        if p.returncode == 0:
            try:
                return json.loads(p.stdout or "null")
            except ValueError:
                raise Failed("unexpected az output") from None
        error = p.stderr or p.stdout
        if attempt < 2 and re.search(r"\b429\b|toomanyrequests|throttl", error, re.IGNORECASE):
            time.sleep(5 * (attempt + 1))
            continue
        raise Failed(reason(error))
    raise Failed("Azure is throttling requests; try again later")


def rest(method: str, url: str, body: Optional[dict] = None, timeout: float = 120):
    """One Azure Resource Manager call. A URL that starts with / goes to the signed-in cloud."""
    args = ["rest", "--method", method, "--url", url]
    if body is not None:
        args += ["--body", json.dumps(body)]
    return az(args, timeout)


def http_json(url: str) -> dict:
    """One GET to the public price list, tried three times when the error may pass."""
    req = urllib.request.Request(url, headers={"User-Agent": f"zoorik-disk-assessment/{VERSION}"})
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code != 429 and e.code < 500:
                break
        except (urllib.error.URLError, http.client.HTTPException, OSError, ValueError):
            pass
        if attempt < 2:
            time.sleep(2 * (attempt + 1))
    raise Failed("the public price list (prices.azure.com) is not reachable")


def parallel(func: Callable, items: list, label: str) -> list:
    """Run func on every item, a few at a time, with a live counter. A failure becomes that item's result.

    On Ctrl-C the items not yet started are dropped, so nothing new starts after it."""
    results: list = [None] * len(items)
    if not items:
        return results

    def guarded(item):
        try:
            return func(item)
        except Failed as e:
            return e
        except Exception as e:  # one odd VM or disk must not stop the whole assessment
            return Failed(f"could not be read ({type(e).__name__})")

    say(f"\r{label}: 0/{len(items)}", end="")
    ex = ThreadPoolExecutor(WORKERS)
    futures = {ex.submit(guarded, item): i for i, item in enumerate(items)}
    try:
        for done, fut in enumerate(as_completed(futures), 1):
            results[futures[fut]] = fut.result()
            say(f"\r{label}: {done}/{len(items)}", end="")
    except BaseException:
        for fut in futures:
            fut.cancel()
        ex.shutdown(wait=False)
        raise
    ex.shutdown()
    say()
    return results


def ask(question: str) -> Optional[bool]:
    """Ask y/N on the terminal. The script itself arrives on stdin, so the answer is read from /dev/tty.

    Keys pressed before the question appears are discarded, so only an answer to it counts.
    Returns None when there is no terminal to ask on."""
    try:
        with open("/dev/tty", encoding="utf-8", errors="replace") as tty:
            try:
                import termios
                termios.tcflush(tty.fileno(), termios.TCIFLUSH)
            except Exception:
                pass
            say(question, end="")
            answer = tty.readline()
    except OSError:
        return None
    if not answer.endswith("\n"):
        say()
    return answer.strip().lower() in ("y", "yes")


# ---------------------------------------------------------------- model

@dataclass
class Subscription:
    id: str
    name: str
    tenant: str
    vms: int = 0
    disks: int = 0
    note: str = ""


@dataclass
class Disk:
    id: str
    name: str
    subscription: str
    resource_group: str
    region: str
    sku: str
    size_gb: int
    tier: str
    state: str
    iops: Optional[int]
    mbps: Optional[int]
    ade: bool
    created: str
    detached: str
    max_shares: int = 1
    price: Optional[float] = None
    price_tier: str = ""
    price_note: str = ""
    hourly: Optional[Dict[str, Tuple[float, float]]] = None
    metric_note: str = ""

    @property
    def kind(self) -> str:
        return DISK_TYPES.get(self.sku.lower(), ("", "", self.sku or "Unknown"))[2]

    @property
    def shared(self) -> bool:
        return self.max_shares > 1

    def peak(self) -> Tuple[Optional[float], Optional[float]]:
        """Peak IOPS and MB/s over the metric window."""
        if not self.hourly:
            return None, None
        return max(v[0] for v in self.hourly.values()), max(v[1] for v in self.hourly.values()) / 1e6


@dataclass
class DataDisk:
    lun: int
    size_gb: int
    disk: Optional[Disk]
    note: str = ""


@dataclass
class Guest:
    kernel: str = ""
    os_id: str = ""
    os_like: str = ""
    os_version: str = ""
    btrfs: str = ""
    luns: Dict[int, List[str]] = field(default_factory=dict)
    mounts: Dict[str, Tuple[str, int]] = field(default_factory=dict)


@dataclass
class Vm:
    id: str
    name: str
    subscription: str
    resource_group: str
    region: str
    zone: str
    size: str
    os_type: str
    power: str
    os_name: str
    os_version: str
    data: List[DataDisk]
    slots: Optional[int] = None
    uncached_iops: int = 0
    uncached_mbps: int = 0
    facts_note: str = ""
    pool_v2: Optional[bool] = None
    managed_by: str = ""
    guest: Optional[Guest] = None
    guest_note: str = ""
    lane: str = ""
    reasons: List[str] = field(default_factory=list)
    used_gb: Optional[float] = None
    util: Optional[float] = None
    current: Optional[float] = None
    pool_gb: Optional[int] = None
    pool_cost: Optional[float] = None
    saving: Optional[float] = None
    no_saving: bool = False

    @property
    def linux(self) -> bool:
        return self.os_type.lower() == "linux"

    @property
    def running(self) -> bool:
        return self.power == "Running"


@dataclass
class RegionFacts:
    sizes: Dict[str, Tuple[int, int, int]]
    v2_zones: Optional[set]
    v2_region: bool = False


@dataclass
class Prices:
    tiers: Dict[str, float] = field(default_factory=dict)
    v2: Dict[str, Dict[float, float]] = field(default_factory=dict)
    ultra: Dict[str, Dict[float, float]] = field(default_factory=dict)


def r2(value: float) -> float:
    return round(value + 1e-9, 2)


def iso(moment: datetime) -> str:
    return moment.strftime("%Y-%m-%dT%H:%M:%SZ")


def parse_time(text: str) -> Optional[datetime]:
    match = re.match(r"(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2}:\d{2})", text or "")
    if not match:
        return None
    return datetime.strptime(match.group(1) + "T" + match.group(2), "%Y-%m-%dT%H:%M:%S").replace(tzinfo=timezone.utc)


# ---------------------------------------------------------------- inventory

SUBSCRIPTIONS_QUERY = "resourcecontainers | where type =~ 'microsoft.resources/subscriptions' | project subscriptionId"
VM_QUERY = """resources
| where type =~ 'microsoft.compute/virtualmachines'
| project id, name, resourceGroup, subscriptionId, location, zones,
    size = tostring(properties.hardwareProfile.vmSize),
    osType = tostring(properties.storageProfile.osDisk.osType),
    power = tostring(properties.extended.instanceView.powerState.code),
    osName = tostring(properties.extended.instanceView.osName),
    osVersion = tostring(properties.extended.instanceView.osVersion),
    dataDisks = properties.storageProfile.dataDisks
| order by id asc"""
DISK_QUERY = """resources
| where type =~ 'microsoft.compute/disks'
| project id, name, resourceGroup, subscriptionId, location,
    sku = tostring(sku.name), tier = tostring(properties.tier),
    sizeGb = toint(properties.diskSizeGB), state = tostring(properties.diskState),
    iops = tolong(properties.diskIOPSReadWrite), mbps = tolong(properties.diskMBpsReadWrite),
    ade = tobool(properties.encryptionSettingsCollection.enabled), maxShares = toint(properties.maxShares),
    created = tostring(properties.timeCreated),
    detached = coalesce(tostring(properties.LastOwnershipUpdateTime), tostring(properties.lastOwnershipUpdateTime))
| order by id asc"""
MANAGED_GROUPS_QUERY = """resourcecontainers
| where type =~ 'microsoft.resources/subscriptions/resourcegroups' and isnotempty(managedBy)
| project subscriptionId, name, managedBy
| order by subscriptionId asc, name asc"""


def signed_in() -> Tuple[str, List[Subscription], List[Subscription]]:
    """Return (directory label, subscriptions to scan, subscriptions left out with the reason)."""
    tenant = (az(["account", "show"]) or {}).get("tenantId", "")
    accounts = [a for a in az(["account", "list"]) or [] if a.get("state") == "Enabled"]
    label = tenant
    for a in accounts:
        if a.get("tenantId") == tenant and a.get("tenantDisplayName"):
            domain = a.get("tenantDefaultDomain")
            label = a["tenantDisplayName"] + (f" ({domain})" if domain else "")
    scan, skipped = [], []
    for a in accounts:
        sub = Subscription(a.get("id", ""), a.get("name", ""), a.get("tenantId", ""))
        if sub.tenant == tenant:
            scan.append(sub)
        else:
            sub.note = (f"in another Azure directory; to include it, run az login --tenant {sub.tenant} "
                        "and run this again")
            skipped.append(sub)
    return label, scan, skipped


def graph(query: str, subscriptions: List[str]) -> list:
    """Run one Resource Graph query over the subscriptions, 1,000 at a time, following every page."""
    rows = []
    for i in range(0, len(subscriptions), 1000):
        options = {"$top": 1000, "resultFormat": "objectArray"}
        while True:
            page = rest("post", GRAPH_URL, {"subscriptions": subscriptions[i:i + 1000], "query": query,
                                            "options": options}) or {}
            rows += page.get("data") or []
            token = page.get("$skipToken")
            if not token:
                break
            options = dict(options, **{"$skipToken": token})
    return rows


def make_disk(r: dict) -> Disk:
    return Disk(id=r.get("id") or "", name=r.get("name") or "", subscription=r.get("subscriptionId") or "",
                resource_group=r.get("resourceGroup") or "", region=(r.get("location") or "").lower(),
                sku=r.get("sku") or "", size_gb=int(r.get("sizeGb") or 0), tier=r.get("tier") or "",
                state=r.get("state") or "", iops=r.get("iops"), mbps=r.get("mbps"), ade=r.get("ade") is True,
                created=r.get("created") or "", detached=r.get("detached") or "",
                max_shares=r.get("maxShares") if isinstance(r.get("maxShares"), int) else 1)


def make_vm(r: dict, disks: Dict[str, Disk]) -> Vm:
    data = []
    for d in r.get("dataDisks") or []:
        if not isinstance(d, dict) or not isinstance(d.get("lun"), int):
            continue
        managed = (d.get("managedDisk") or {}).get("id") or ""
        disk = disks.get(managed.lower()) if managed else None
        note = "" if disk else "unmanaged (VHD) disk; not priced" if not managed else "disk details not readable"
        data.append(DataDisk(lun=d["lun"], size_gb=int((disk.size_gb if disk else d.get("diskSizeGB")) or 0),
                             disk=disk, note=note))
    power = (r.get("power") or "").split("/")[-1]
    return Vm(id=r.get("id") or "", name=r.get("name") or "", subscription=r.get("subscriptionId") or "",
              resource_group=r.get("resourceGroup") or "", region=(r.get("location") or "").lower(),
              zone=",".join(r.get("zones") or []), size=r.get("size") or "", os_type=r.get("osType") or "",
              power=power.capitalize(), os_name=r.get("osName") or "", os_version=r.get("osVersion") or "",
              data=sorted(data, key=lambda x: x.lun))


def managing_service(managed_by: str) -> str:
    """The resource type that manages a resource group, e.g. Microsoft.Databricks/workspaces, without its name."""
    match = re.search(r"/providers/([^/]+/[^/]+)", managed_by or "")
    return match.group(1) if match else "another Azure service"


def inventory(subs: List[Subscription]) -> Tuple[List[Vm], List[Disk]]:
    """Every VM and managed disk in the subscriptions, through Azure Resource Graph."""
    if not subs:
        return [], []
    try:
        seen = {r.get("subscriptionId") for r in graph(SUBSCRIPTIONS_QUERY, [s.id for s in subs])}
    except Failed as e:
        for s in subs:
            s.note = f"Resource Graph: {e}"
        return [], []
    for s in subs:
        if s.id not in seen:
            s.note = "Resource Graph cannot read it (no Reader access?)"
    readable = [s.id for s in subs if not s.note]
    if not readable:
        return [], []
    disk_list = [make_disk(r) for r in graph(DISK_QUERY, readable)]
    disks = {d.id.lower(): d for d in disk_list}
    vms = [make_vm(r, disks) for r in graph(VM_QUERY, readable)]
    try:
        managed = {(r.get("subscriptionId"), (r.get("name") or "").lower()): managing_service(r.get("managedBy"))
                   for r in graph(MANAGED_GROUPS_QUERY, readable)}
    except Failed:
        managed = {}
    for v in vms:
        v.managed_by = managed.get((v.subscription, v.resource_group.lower()), "")
    for s in subs:
        s.vms = sum(1 for v in vms if v.subscription == s.id)
        s.disks = sum(1 for d in disk_list if d.subscription == s.id)
    return vms, disk_list


# ---------------------------------------------------------------- VM sizes and prices

def region_facts(key: Tuple[str, str]) -> RegionFacts:
    """VM size limits in a region, and where in it Premium SSD v2 is offered to this subscription."""
    sub, region = key
    url = (f"/subscriptions/{sub}/providers/Microsoft.Compute/skus?api-version={SKUS_API}&%24filter="
           + urllib.parse.quote(f"location eq '{region}'"))
    facts = RegionFacts(sizes={}, v2_zones=None)
    items = []
    while url:
        page = rest("get", url, timeout=300) or {}
        items += page.get("value") or []
        url = page.get("nextLink") or ""
    for item in items:
        caps = {c.get("name"): str(c.get("value") or "") for c in item.get("capabilities") or []}
        if item.get("resourceType") == "virtualMachines" and caps.get("MaxDataDiskCount", "").isdigit():
            iops, bps = caps.get("UncachedDiskIOPS", ""), caps.get("UncachedDiskBytesPerSecond", "")
            facts.sizes[(item.get("name") or "").lower()] = (int(caps["MaxDataDiskCount"]),
                                                             int(iops) if iops.isdigit() else 0,
                                                             int(bps) // 1000000 if bps.isdigit() else 0)
        elif item.get("resourceType") == "disks" and item.get("name") == "PremiumV2_LRS":
            info = (item.get("locationInfo") or [{}])[0]
            zones, offered = set(info.get("zones") or []), True
            for r in item.get("restrictions") or []:
                if r.get("type") == "Location":
                    zones, offered = set(), False
                elif r.get("type") == "Zone":
                    zones -= set((r.get("restrictionInfo") or {}).get("zones") or [])
            facts.v2_zones = (facts.v2_zones or set()) | zones
            facts.v2_region = facts.v2_region or offered
    return facts


def v2_offered(vm: Vm, facts: RegionFacts) -> bool:
    """Whether the VM can take Premium SSD v2 disks: in its zone, or, without a zone, in its region."""
    if vm.zone:
        return vm.zone in (facts.v2_zones or set())
    return facts.v2_region and vm.region in V2_NONZONAL_REGIONS


def region_prices(region: str) -> Prices:
    """List prices of every managed-disk meter in one region, from the public Azure Retail Prices API."""
    products = " or ".join(f"productName eq '{p}'" for p in TIER_PRODUCTS + tuple(PROVISIONED_PRODUCTS))
    query = f"armRegionName eq '{region}' and priceType eq 'Consumption' and ({products})"
    url = PRICES_URL + "?" + urllib.parse.urlencode({"$filter": query})
    prices = Prices()
    while url:
        page = http_json(url)
        for item in page.get("Items") or []:
            add_price(prices, item)
        url = page.get("NextPageLink")
    return prices


def add_price(prices: Prices, item: dict) -> None:
    product, meter, unit = item.get("productName", ""), item.get("meterName", ""), item.get("unitOfMeasure", "")
    price = item.get("unitPrice")
    if not isinstance(price, (int, float)):
        return
    tier = re.fullmatch(r"([PES]\d+) (LRS|ZRS) Disk", meter)
    if tier and product in TIER_PRODUCTS and unit == "1/Month":
        prices.tiers[f"{tier.group(1)} {tier.group(2)}"] = float(price)
        return
    prefix = PROVISIONED_PRODUCTS.get(product)
    if prefix and meter.startswith(prefix) and unit in ("1 GiB/Hour", "1/Hour"):
        kind = PROVISIONED_METERS.get(meter[len(prefix):])
        if kind:
            store = prices.v2 if product == "Azure Premium SSD v2" else prices.ultra
            store.setdefault(kind, {})[float(item.get("tierMinimumUnits") or 0)] = float(price)


def graduated(units: float, tiers: Dict[float, float]) -> float:
    """Price of units on a tiered meter: each tier's price applies from its first unit to the next tier's."""
    steps = sorted(tiers.items())
    total = 0.0
    for i, (start, price) in enumerate(steps):
        end = steps[i + 1][0] if i + 1 < len(steps) else float("inf")
        if units > start:
            total += (min(units, end) - start) * price
    return total


def provisioned_monthly(meters: Dict[str, Dict[float, float]], gib: float, iops: float, mbps: float) -> Optional[float]:
    """$/month of a disk billed by provisioned GiB, IOPS and MB/s (Premium SSD v2, Ultra Disk)."""
    if not all(meters.get(k) for k in ("capacity", "iops", "mbps")):
        return None
    hourly = (graduated(gib, meters["capacity"]) + graduated(iops, meters["iops"])
              + graduated(mbps, meters["mbps"]))
    return r2(hourly * HOURS_PER_MONTH)


def billed_tier(letter: str, size_gb: int) -> Optional[str]:
    """The tier Azure bills a disk at: the smallest offered size at or above the disk's size."""
    for limit, number in (HDD_TIER_SIZES if letter == "S" else TIER_SIZES):
        if size_gb <= limit:
            return letter + number
    return None


def price_disk(d: Disk, prices: Dict[str, Optional[Prices]]) -> None:
    """Set the disk's list price per month, or the reason it has none."""
    family, redundancy, label = DISK_TYPES.get(d.sku.lower(), ("", "", d.sku))
    region_price = prices.get(d.region)
    if not family:
        d.price_note = f"{d.sku or 'unknown'} disks are not priced"
        return
    if region_price is None:
        d.price_note = "price list not readable for " + d.region
        return
    if family in ("V2", "ULTRA"):
        if d.iops is None or d.mbps is None:
            d.price_note = "provisioned IOPS or MB/s not reported"
            return
        meters = region_price.v2 if family == "V2" else region_price.ultra
        d.price = provisioned_monthly(meters, d.size_gb, d.iops, d.mbps)
        d.price_note = "" if d.price is not None else f"no {label} list price in {d.region}"
        return
    tier = billed_tier(family, d.size_gb)
    names = [t[1] for t in (HDD_TIER_SIZES if family == "S" else TIER_SIZES)]
    if tier and family == "P" and re.fullmatch(r"P\d+", d.tier or "") and d.tier[1:] in names \
            and names.index(d.tier[1:]) > names.index(tier[1:]):
        tier = d.tier
    if not tier:
        d.price_note = f"{d.size_gb:,} GB is above the largest {label} tier"
        return
    d.price_tier = tier
    monthly = region_price.tiers.get(f"{tier} {redundancy}")
    d.price = None if monthly is None else r2(monthly)
    if d.price is None:
        d.price_note = f"no list price for {tier} {redundancy} in {d.region}"


# ---------------------------------------------------------------- measured performance

def disk_metrics(disk: Disk) -> Dict[str, Tuple[float, float]]:
    """IOPS and bytes/s of one disk for each hour of the last 14 days: the busiest minute of the hour (Maximum),
    reads plus writes, or the hour's average where Azure Monitor gives no maximum."""
    end = datetime.now(timezone.utc).replace(minute=0, second=0, microsecond=0)
    start = end - timedelta(days=METRIC_DAYS)
    query = urllib.parse.urlencode({"api-version": METRICS_API, "metricnames": ",".join(METRICS),
                                    "timespan": f"{iso(start)}/{iso(end)}", "interval": "PT1H",
                                    "aggregation": "Average,Maximum"}, quote_via=urllib.parse.quote)
    data = rest("get", f"{disk.id}/providers/Microsoft.Insights/metrics?{query}") or {}
    hours: Dict[str, Tuple[float, float]] = {}
    for metric in data.get("value") or []:
        is_ops = "Operations" in ((metric.get("name") or {}).get("value") or "")
        for series in metric.get("timeseries") or []:
            for point in series.get("data") or []:
                value, stamp = point.get("maximum"), point.get("timeStamp")
                if not isinstance(value, (int, float)):
                    value = point.get("average")
                if not isinstance(value, (int, float)) or not stamp:
                    continue
                ops, rate = hours.get(stamp, (0.0, 0.0))
                hours[stamp] = (ops + value, rate) if is_ops else (ops, rate + value)
    return hours


# ---------------------------------------------------------------- inside the VM

RUN_COMMAND_STARTED: List[str] = []


def read_guest(vm: Vm) -> Guest:
    """Run the command that only reads in the VM through Azure Run Command and parse what it printed."""
    RUN_COMMAND_STARTED.append(vm.id)
    out = az(["vm", "run-command", "invoke", "--ids", vm.id, "--command-id", "RunShellScript",
              "--scripts", GUEST_SCRIPT], RUN_COMMAND_TIMEOUT) or {}
    message = "".join(str((v or {}).get("message") or "") for v in out.get("value") or [])
    if "[stdout]" in message:
        message = message.split("[stdout]", 1)[1].split("[stderr]", 1)[0]
    return parse_guest(message)


def unescape_lsblk(text: str) -> str:
    return re.sub(r"\\x([0-9a-fA-F]{2})", lambda m: chr(int(m.group(1), 16)), text)


def parse_guest(text: str) -> Guest:
    lines = [line.strip("\r") for line in text.splitlines()]
    if "S" not in lines or "E" not in lines:
        raise Failed("the command's output came back incomplete")
    g = Guest()
    for line in lines:
        key, _, value = line.partition("=")
        if key in ("ID", "ID_LIKE", "VERSION_ID") and value:
            value = value.strip().strip("\"'")
            if key == "ID":
                g.os_id = value.lower()
            elif key == "ID_LIKE":
                g.os_like = value.lower()
            else:
                g.os_version = value
            continue
        tag, _, rest_of_line = line.partition(" ")
        if tag == "K":
            g.kernel = rest_of_line.strip()
        elif tag == "B":
            g.btrfs = rest_of_line.strip()
        elif tag == "L":
            lun, _, mounts = rest_of_line.partition(" ")
            if lun.isdigit():
                g.luns[int(lun)] = [unescape_lsblk(m) for m in mounts.split() if m.startswith("/")]
        elif tag == "D":
            parts = rest_of_line.split(" ", 3)
            if len(parts) == 4 and parts[1].isdigit() and parts[2].isdigit():
                g.mounts[parts[3]] = (parts[0], int(parts[2]))
    if not g.kernel or g.btrfs not in ("loaded", "module", "absent"):
        raise Failed("the command's output was not readable")
    return g


def kernel_version(kernel: str) -> Optional[Tuple[int, int]]:
    match = re.match(r"(\d+)\.(\d+)", kernel or "")
    return (int(match.group(1)), int(match.group(2))) if match else None


def rhel_family(vm: Vm) -> bool:
    g = vm.guest
    if g and g.os_id:
        return g.os_id in RHEL_FAMILY or "rhel" in g.os_like.split()
    name = vm.os_name.lower()
    return any(word in name for word in ("rhel", "red hat", "redhat", "centos", "rocky", "alma", "oracle"))


def os_label(vm: Vm) -> str:
    g = vm.guest
    if g and g.os_id:
        return f"{OS_NAMES.get(g.os_id, g.os_id.capitalize())} {g.os_version}".strip()
    if vm.os_name:
        return f"{OS_NAMES.get(vm.os_name.lower(), vm.os_name)} {vm.os_version}".strip()
    return vm.os_type.capitalize()


# ---------------------------------------------------------------- pool sizing (as Zoorik ElasticVol plans it)

def ceil_div(numerator: int, divisor: int) -> int:
    return numerator if divisor <= 0 else -(-numerator // divisor)


def half_up(value: float) -> int:
    return int(math.floor(value + 0.5))


def pool_size_gib(used_gib: int) -> int:
    """The pool holds today's data plus headroom: 20% of it, at least 50 GB."""
    return math.ceil(used_gib + max(HEADROOM_FRACTION * used_gib, HEADROOM_MIN_GIB))


def tier_up(gib: int, v2: bool) -> int:
    """A Premium SSD v2 disk takes any size; a Premium SSD disk the smallest Premium SSD size at or above it."""
    return gib if v2 else next((t for t in V1_TIER_GIB if t >= gib), V1_TIER_GIB[-1])


def tier_down(gib: int, v2: bool) -> int:
    """A Premium SSD v2 disk takes any size; a Premium SSD disk the largest Premium SSD size at or below it."""
    return gib if v2 else max((t for t in V1_TIER_GIB if t <= gib), default=V1_TIER_GIB[0])


def elastic_floor(v2: bool) -> int:
    return ELASTIC_FLOOR_V2_GIB if v2 else ELASTIC_FLOOR_V1_GIB


def elastic_cap_gib(uncached_mbps: int, floor: int) -> int:
    """The largest elastic disk the VM can empty in the evacuation window."""
    if uncached_mbps <= 0:
        return ELASTIC_CAP_GIB
    evacuable = EVAC_WINDOW_MINUTES * 60.0 * EVAC_SHARE * uncached_mbps / (2.0 * ELASTIC_FILL_FRACTION * 1024.0)
    return max(floor, min(ELASTIC_CAP_GIB, math.floor(evacuable)))


def clamp_elastics(count: int) -> int:
    return max(MIN_ELASTICS, min(MAX_ELASTICS, count))


def full_geometry(pool: int, used: int, uncached_mbps: int, v2: bool) -> List[int]:
    """Two anchors and three to six equal elastics (data member sizes): the shape when slots allow."""
    floor = elastic_floor(v2)
    quantum = tier_up(max(floor, min(elastic_cap_gib(uncached_mbps, floor), pool // ELASTIC_DIVISOR)), v2)
    if pool < ANCHOR_POOL_THRESHOLD_GIB:
        quantum = max(quantum, tier_up(ceil_div(pool, MAX_ELASTICS), v2))
        return [quantum] * clamp_elastics(ceil_div(pool, quantum))
    by_band = half_up(pool * ANCHOR_CAP_FRACTION)
    joint = min(by_band, used) if used > 0 else by_band
    anchor = tier_down(max(1, min(joint // ANCHOR_COUNT, ANCHOR_MAX_GIB)), v2)
    remaining = max(0, pool - anchor * ANCHOR_COUNT)
    return [anchor] * ANCHOR_COUNT + [quantum] * clamp_elastics(ceil_div(remaining, quantum))


def covering_quantum(span: int, count: int, v2: bool) -> int:
    floor = elastic_floor(v2)
    return floor if span <= 0 else max(floor, tier_up(ceil_div(span, count), v2))


def degraded_geometry(pool: int, uncached_mbps: int, member_slots: int,
                      v2: bool) -> Optional[Tuple[List[int], int]]:
    """One anchor and up to three elastics: the shape when the VM has fewer free slots."""
    small = 1 if v2 else 0
    elastics = min(DEGRADED_MAX_ELASTICS, member_slots - small - DEGRADED_ANCHORS)
    if elastics < 1:
        small = 0
        elastics = min(DEGRADED_MAX_ELASTICS, member_slots - DEGRADED_ANCHORS)
    if elastics < 1:
        return None
    ceiling = tier_down(elastic_cap_gib(uncached_mbps, elastic_floor(v2)), v2)
    anchor = tier_down(min(half_up(pool * ANCHOR_CAP_FRACTION), ANCHOR_MAX_GIB), v2)
    quantum = covering_quantum(pool - anchor, elastics, v2)
    if quantum > ceiling:
        anchor = min(ANCHOR_MAX_GIB, tier_up(max(1, pool - elastics * ceiling), v2))
        quantum = covering_quantum(pool - anchor, elastics, v2)
    if anchor <= 0 or quantum <= 0 or quantum > ceiling or anchor + quantum * elastics < pool:
        return None
    return [anchor] + [quantum] * elastics, small


def pool_members(pool: int, used: int, slots: int, attached: int, uncached_mbps: int,
                 v2: bool) -> Optional[Tuple[List[int], int]]:
    """(data member sizes in GiB, number of small members), or None when the VM cannot host the pool.

    The pool is built beside today's disks, so its members must fit the slots those leave free. Only a
    Premium SSD v2 pool has the small member."""
    full = full_geometry(pool, used, uncached_mbps, v2)
    small = 1 if v2 else 0
    member_slots = max(0, slots - max(0, attached) - RESERVED_SLOTS)
    if len(full) + small <= member_slots:
        return full, small
    if slots < MIN_VM_SLOTS or member_slots < MIN_POOL_SLOTS:
        return None
    return degraded_geometry(pool, uncached_mbps, member_slots, v2)


def member_iops(demand: float, members: int, size: int, vm_iops: int) -> int:
    count = max(1, members)
    iops = max(V2_MIN_IOPS, math.ceil(DEMAND_FACTOR * max(0.0, demand) / count))
    iops = min(iops, max(V2_MIN_IOPS, size * V2_MAX_IOPS_PER_GIB))
    if vm_iops > 0:
        iops = min(iops, max(V2_MIN_IOPS, vm_iops // count))
    return iops


def member_mbps(demand: float, members: int, iops: int, vm_mbps: int) -> int:
    count = max(1, members)
    mbps = max(V2_MIN_MBPS, math.ceil(DEMAND_FACTOR * max(0.0, demand) / count))
    mbps = min(mbps, max(V2_MIN_MBPS, min(V2_MAX_MBPS, int(iops * V2_MBPS_PER_IOPS))))
    if vm_mbps > 0:
        mbps = min(mbps, max(V2_MIN_MBPS, vm_mbps // count))
    return mbps


# ---------------------------------------------------------------- assessment

def screen(vm: Vm) -> List[str]:
    """Reasons ElasticVol cannot take the VM, from what Azure reports (before looking inside it)."""
    red = []
    if not vm.linux:
        red.append(f"{vm.os_type or 'Non-Linux'} VM; ElasticVol supports Linux")
    if vm.managed_by:
        red.append(f"in a resource group managed by {vm.managed_by}; ElasticVol does not take VMs another Azure "
                   "service runs")
    if vm.slots is not None and vm.slots < MIN_VM_SLOTS:
        red.append(f"{vm.size} has {vm.slots} data-disk slots; ElasticVol needs {MIN_VM_SLOTS} or more")
    if vm.linux and rhel_family(vm):
        red.append("RHEL-family Linux; its stock kernel has no btrfs")
    return red


def data_mounts(vm: Vm) -> Dict[str, List[int]]:
    """Mounts that sit on the VM's data disks, with the LUNs under each."""
    mounts: Dict[str, List[int]] = {}
    known = {dd.lun for dd in vm.data}
    for lun, points in (vm.guest.luns if vm.guest else {}).items():
        if lun in known:
            for point in points:
                mounts.setdefault(point, []).append(lun)
    return mounts


def pooled_peak(pooled: List[DataDisk]) -> Tuple[float, float]:
    """The measured peak of the disks the pool replaces, summed across them per hour: (IOPS, MB/s)."""
    hours: Dict[str, Tuple[float, float]] = {}
    for dd in pooled:
        for stamp, (ops, rate) in dd.disk.hourly.items():
            total = hours.get(stamp, (0.0, 0.0))
            hours[stamp] = (total[0] + ops, total[1] + rate)
    return max(v[0] for v in hours.values()), max(v[1] for v in hours.values()) / 1e6


def pool_cost_v2(region_price: Prices, vm: Vm, members: List[int], peak_iops: float, peak_mbps: float) -> float:
    """$/month of a Premium SSD v2 pool: each disk's size, plus IOPS and MB/s above the included baseline
    where the measured peak needs them."""
    cost = 0.0
    for size in members:
        iops = member_iops(peak_iops, len(members), size, vm.uncached_iops)
        price = provisioned_monthly(region_price.v2, size, iops,
                                    member_mbps(peak_mbps, len(members), iops, vm.uncached_mbps))
        if price is None:
            raise Failed(f"no Premium SSD v2 list price in {vm.region}")
        cost += price
    return cost


def pool_cost_v1(region_price: Prices, vm: Vm, members: List[int]) -> float:
    """$/month of a Premium SSD pool: each disk at the list price of its Premium SSD tier."""
    cost = 0.0
    for size in members:
        tier = billed_tier("P", size)
        monthly = region_price.tiers.get(f"{tier} LRS")
        if monthly is None:
            raise Failed(f"no Premium SSD list price for {tier} LRS in {vm.region}")
        cost += r2(monthly)
    return cost


def v1_shortfall(vm: Vm, members: List[int], peak_iops: float, peak_mbps: float) -> str:
    """Why a Premium SSD pool cannot carry the measured peak x 1.3 (or what the VM can drive, if less); else ''."""
    need_iops, need_mbps = DEMAND_FACTOR * peak_iops, DEMAND_FACTOR * peak_mbps
    if vm.uncached_iops > 0:
        need_iops = min(need_iops, vm.uncached_iops)
    if vm.uncached_mbps > 0:
        need_mbps = min(need_mbps, vm.uncached_mbps)
    has_iops = sum(V1_TIER_PERFORMANCE[tier_up(size, False)][0] for size in members)
    has_mbps = sum(V1_TIER_PERFORMANCE[tier_up(size, False)][1] for size in members)
    if need_iops <= has_iops and need_mbps <= has_mbps:
        return ""
    return (f"not estimated: a Premium SSD pool of this size carries {has_iops:,} IOPS and {has_mbps:,} MB/s, "
            f"less than the measured peak x 1.3 ({math.ceil(need_iops):,} IOPS, {math.ceil(need_mbps):,} MB/s)")


def assess(vm: Vm, prices: Dict[str, Optional[Prices]]) -> None:
    """Lane, used space and the saving for one VM; every blank figure gets its reason."""
    red, amber, notes = screen(vm), [], []
    g, unread = vm.guest, vm.guest_note
    if g and not g.luns:
        g = None
        unread = vm.guest_note = "the VM shows no /dev/disk/azure/scsi1 links to match its data disks"
    shared = {dd.lun for dd in vm.data if dd.disk and dd.disk.shared}
    on_data = data_mounts(vm) if g else {}
    mounts = {m: luns for m, luns in on_data.items() if not shared & set(luns)}
    if g:
        version = kernel_version(g.kernel)
        if not rhel_family(vm) and g.btrfs not in ("loaded", "module"):
            red.append("btrfs is not available in this kernel")
        if version and version < KERNEL_FLOOR:
            amber.append(f"kernel {g.kernel} is below 5.14; upgrade the kernel first (Ubuntu 20.04: the HWE kernel)")
        if not mounts:
            red.append("its mounted data disks are shared disks, which ElasticVol does not pool" if on_data
                       else "no mounted data disk found inside the VM")
    if any(dd.disk and dd.disk.ade for dd in vm.data):
        amber.append("Azure Disk Encryption is on; move to server-side encryption first")

    unpriced = [dd for dd in vm.data if not dd.disk or dd.disk.price is None]
    vm.current = None if unpriced else r2(sum(dd.disk.price for dd in vm.data))
    if unpriced:
        first = unpriced[0]
        notes.append(f"today's cost unknown: {first.note or first.disk.price_note}")

    pooled = sorted({lun for luns in mounts.values() for lun in luns})
    pooled_disks = [dd for dd in vm.data if dd.lun in pooled]
    missing = [m for m in mounts if m not in g.mounts] if g else []
    if g and mounts and not missing:
        used_bytes = sum(g.mounts[m][1] for m in mounts)
        vm.used_gb = used_bytes / GIB
        provisioned = sum(dd.size_gb for dd in pooled_disks)
        vm.util = vm.used_gb / provisioned if provisioned else None
    elif missing:
        notes.append(f"df did not report {missing[0]}")
    if g and mounts:
        if shared:
            notes.append(f"{len(shared)} shared disk(s) stay as they are; ElasticVol does not pool a shared disk")
        left = len([dd for dd in vm.data if dd.lun not in pooled and dd.lun not in shared])
        if left:
            notes.append(f"{left} data disk(s) without a mounted file system stay as they are")

    def finish(lane: str, why: str = "") -> None:
        vm.lane = lane
        vm.reasons = red + amber + ([why] if why else []) + notes

    if red:
        return finish("Not supported")
    if not g:
        return finish("Not checked", f"disk usage not read: {unread}" if unread else "")
    lane = "Needs work" if amber else "Ready"
    if missing or vm.used_gb is None:
        return finish(lane)
    if vm.slots is None:
        return finish(lane, f"VM size limits not readable: {vm.facts_note or 'size not listed'}")
    if any(dd.disk is None or dd.disk.price is None for dd in pooled_disks):
        return finish(lane)
    if any(dd.disk.hourly is None for dd in pooled_disks):
        bad = next(dd.disk for dd in pooled_disks if dd.disk.hourly is None)
        return finish(lane, f"IOPS not measured: {bad.metric_note or 'no metric data'}")

    used = sum(max(1, math.ceil(g.mounts[m][1] / GIB)) for m in mounts)
    pool = pool_size_gib(used)
    layout = pool_members(pool, used, vm.slots, len(vm.data), vm.uncached_mbps, vm.pool_v2)
    if layout is None:
        amber.append("not enough free data-disk slots to build the pool beside today's disks")
        return finish("Needs work")
    data_members, small = layout
    if sum(data_members) < pool:
        return finish(lane, f"{used:,} GB used is more than one pool holds; not estimated")
    members = data_members + [SMALL_MEMBER_GIB] * small
    peak_iops, peak_mbps = pooled_peak(pooled_disks)
    if not vm.pool_v2:
        short = v1_shortfall(vm, members, peak_iops, peak_mbps)
        if short:
            return finish(lane, short)
    region_price = prices.get(vm.region) or Prices()
    try:
        cost = (pool_cost_v2(region_price, vm, members, peak_iops, peak_mbps) if vm.pool_v2
                else pool_cost_v1(region_price, vm, members))
    except Failed as e:
        return finish(lane, str(e))
    vm.pool_gb = sum(members)
    vm.pool_cost = r2(cost)
    today = r2(sum(dd.disk.price for dd in pooled_disks))
    if today - vm.pool_cost <= 0:
        vm.no_saving = True
        return finish(lane, f"the pool (${vm.pool_cost:,.2f}/mo) would cost no less than these disks "
                            f"(${today:,.2f}/mo)")
    vm.saving = r2(today - vm.pool_cost)
    if any(dd.disk.sku.lower().endswith("_zrs") for dd in pooled_disks):
        notes.append("today's disks are zone-redundant (ZRS); the pool is locally redundant (LRS), and part of "
                     "the saving is that difference")
    finish(lane)


# ---------------------------------------------------------------- report

def money(v: Optional[float]) -> str:
    return "not available" if v is None else f"${v:,.2f}"


def gb(v: Optional[float]) -> Optional[int]:
    return None if v is None else int(round(v))


def attached_once(vms: List[Vm]) -> List[DataDisk]:
    """The VMs' data disks, each once: a shared disk shows on every VM it is attached to but is billed once."""
    seen, out = set(), []
    for v in vms:
        for dd in v.data:
            key = dd.disk.id.lower() if dd.disk else (v.id, dd.lun)
            if key not in seen:
                seen.add(key)
                out.append(dd)
    return out


def priced(disks: List[DataDisk]) -> List[float]:
    return [dd.disk.price for dd in disks if dd.disk and dd.disk.price is not None]


def no_saving_text(vms: List[Vm]) -> str:
    """What the saving figure says when no VM has a saving."""
    read = [v for v in vms if v.data and v.linux and v.guest and v.guest.luns]
    estimated = [v for v in vms if v.saving or v.no_saving]
    if not read:
        return "not estimated: no VM had its disk usage read"
    if estimated:
        return f"none found on the {len(estimated)} VM(s) estimated"
    return "not estimated on the VMs that were read; see the Reason column on the VMs sheet"


VM_COLS = [("Subscription", "text"), ("Resource group", "text"), ("VM", "text"), ("Region", "text"),
           ("Zone", "text"), ("Size", "text"), ("OS", "text"), ("Power state", "text"),
           ("Data-disk slots", "int"), ("Data disks", "int"), ("Provisioned GB", "int"), ("Used GB", "int"),
           ("Mounted-disk utilization %", "pct"), ("Current $/mo", "usd"), ("Pool disks", "text"),
           ("Pool GB", "int"), ("Pool $/mo", "usd"), ("Saving $/mo", "usd"), ("Lane", "text"), ("Reason", "text")]
DISK_COLS = [("VM", "text"), ("Resource group", "text"), ("Disk", "text"), ("LUN", "int"), ("Disk type", "text"),
             ("Tier", "text"), ("GB", "int"), ("Provisioned IOPS", "int"), ("Provisioned MB/s", "int"),
             ("Peak IOPS (14 days)", "int"), ("Peak MB/s (14 days)", "int"), ("Mount", "text"), ("Used GB", "int"),
             ("$/mo", "usd"), ("Note", "text")]
UNATTACHED_COLS = [("Subscription", "text"), ("Resource group", "text"), ("Disk", "text"), ("Region", "text"),
                   ("Disk type", "text"), ("Tier", "text"), ("GB", "int"), ("$/mo", "usd"),
                   ("Days since detached", "int"), ("Note", "text")]


def pool_disks(vm: Vm) -> str:
    """The disk type the VM's pool would use; blank where ElasticVol cannot take the VM or it is not known."""
    if not vm.linux or vm.lane in ("", "Not supported") or vm.pool_v2 is None:
        return ""
    return POOL_V2 if vm.pool_v2 else POOL_V1


def vm_rows(vms: List[Vm], names: Dict[str, str]) -> Tuple[list, list]:
    listed = sorted((v for v in vms if v.data), key=lambda v: (-(v.saving or 0), -(v.current or 0), v.name))
    rows = []
    for v in listed:
        saving = "No saving" if v.no_saving else v.saving
        rows.append([names.get(v.subscription, v.subscription), v.resource_group, v.name, v.region, v.zone, v.size,
                     os_label(v), v.power, v.slots, len(v.data), sum(dd.size_gb for dd in v.data), gb(v.used_gb),
                     v.util, v.current, pool_disks(v), v.pool_gb, v.pool_cost, saving, v.lane, "; ".join(v.reasons)])
    once = attached_once(listed)
    spend = priced(once)
    notes = []
    if spend and any(v.current is None for v in listed):
        notes.append("Current $/mo adds every priced disk, also on VMs whose own cell is blank")
    if len(once) < sum(len(v.data) for v in listed):
        notes.append("a shared disk counts once")
    total = ["TOTAL", "", "", "", "", "", "", "", None, len(once), sum(dd.size_gb for dd in once), None, None,
             r2(sum(spend)) if spend else None, "", None, None,
             r2(sum(v.saving for v in listed if v.saving)) if any(v.saving for v in listed) else None, "",
             "; ".join(notes)]
    return rows, total


def disk_rows(vms: List[Vm]) -> list:
    rows = []
    for v in sorted(vms, key=lambda v: (v.name, v.resource_group)):
        mounts = data_mounts(v)
        for dd in v.data:
            d = dd.disk
            if d is None:
                continue
            on_disk = sorted(m for m, luns in mounts.items() if dd.lun in luns)
            used, notes = None, [d.price_note] if d.price_note else []
            if d.shared:
                notes.append("shared disk; not pooled")
            if on_disk and v.guest and all(m in v.guest.mounts for m in on_disk):
                if all(len(mounts[m]) == 1 for m in on_disk):
                    used = gb(sum(v.guest.mounts[m][1] for m in on_disk) / GIB)
                else:
                    notes.append("mount spans several disks; used space is on the VM row")
            peak_iops, peak_mbps = d.peak()
            if d.hourly is None:
                notes.append(d.metric_note)
            rows.append([v.name, v.resource_group, d.name, dd.lun, d.kind, d.price_tier, d.size_gb, d.iops, d.mbps,
                         gb(peak_iops), gb(peak_mbps), ", ".join(on_disk), used, d.price,
                         "; ".join(n for n in notes if n)])
    return rows


def unattached_rows(disks: List[Disk], names: Dict[str, str], now: datetime) -> Tuple[list, list]:
    rows = []
    for d in sorted((d for d in disks if d.state.lower() == "unattached"), key=lambda d: (-(d.price or 0), d.name)):
        detached, created, note = parse_time(d.detached), parse_time(d.created), d.price_note
        days = (now - detached).days if detached else None
        if days is None:
            note = "; ".join(x for x in (note, "Azure did not record when it was detached"
                                         + (f"; created {created:%Y-%m-%d}" if created else "")) if x)
        rows.append([names.get(d.subscription, d.subscription), d.resource_group, d.name, d.region, d.kind,
                     d.price_tier, d.size_gb, d.price, days, note])
    known = [r[7] for r in rows if r[7] is not None]
    unknown = len(rows) - len(known)
    total = ["TOTAL", "", "", "", "", "", sum(r[6] for r in rows), r2(sum(known)) if known else None, None,
             f"{unknown} disk(s) not priced" if unknown else ""]
    return rows, total


def summary_rows(run: dict, vms: List[Vm], disks: List[Disk]) -> Tuple[list, set]:
    """Two columns, Item and Value, in titled sections."""
    rows, bold = [], set()

    def section(title: str) -> None:
        bold.add(len(rows))
        rows.append([title, ""])

    def item(label: str, value, fmt: str = "text") -> None:
        rows.append([label, (value, fmt)])

    with_data = [v for v in vms if v.data]
    linux = [v for v in with_data if v.linux]
    attached = attached_once(with_data)
    spend = priced(attached)
    linux_spend = priced(attached_once(linux))
    savings = [v for v in with_data if v.saving]
    total_saving = r2(sum(v.saving for v in savings)) if savings else None
    unestimated = [v for v in linux if v.saving is None and not v.no_saving and v.lane != "Not supported"]
    scanned = [s for s in run["scanned"] if not s.note]

    section("Run")
    item("Directory (tenant)", run["directory"])
    item("Subscriptions scanned", len(scanned), "int")
    item("Subscriptions with VMs or disks", ", ".join(s.name for s in scanned if s.vms or s.disks) or "none")
    for s in [s for s in run["scanned"] if s.note] + run["skipped"]:
        item(f"Not scanned: {s.name} ({s.id})", s.note)
    item("Run on", run["when"])
    item("Script version", VERSION)

    section("Estate")
    item("VMs", len(vms), "int")
    item("VMs running", sum(1 for v in vms if v.running), "int")
    item("VMs with data disks", len(with_data), "int")
    item("Data disks attached", len(attached), "int")
    item("Data-disk capacity, GB", sum(dd.size_gb for dd in attached), "int")
    item("Data-disk spend today, $/mo", r2(sum(spend)) if spend else None, "usd")
    item("of which on Linux VMs, $/mo", r2(sum(linux_spend)) if linux_spend else None, "usd")
    unpriced = len(attached) - len(spend)
    if unpriced:
        item("Data disks not priced", f"{unpriced} (see the Reason column on the VMs sheet)")

    section("Zoorik ElasticVol")
    if total_saving is None:
        item("Estimated saving, $/mo", no_saving_text(vms))
    else:
        item("Estimated saving, $/mo", total_saving, "usd")
        item("VMs with a saving", len(savings), "int")
        item("Saving as share of data-disk spend", total_saving / sum(spend) if spend else None, "pct")
    unestimated_spend = priced(attached_once([v for v in unestimated if v.current is not None]))
    if unestimated_spend:
        item("Spend on Linux VMs not yet estimated, $/mo", r2(sum(unestimated_spend)), "usd")
    for lane in ("Ready", "Needs work", "Not supported", "Not checked"):
        count = sum(1 for v in with_data if v.lane == lane)
        if count:
            item(f"Lane: {lane}", count, "int")
    item("VMs with disk usage read", sum(1 for v in linux if v.guest and v.guest.luns), "int")
    why_not: Dict[str, int] = {}
    for v in linux:
        if not (v.guest and v.guest.luns) and v.guest_note:
            why_not[v.guest_note] = why_not.get(v.guest_note, 0) + 1
    for note, count in sorted(why_not.items(), key=lambda kv: -kv[1]):
        item(f"Not read: {note}", count, "int")

    section("Unattached disks (not part of the ElasticVol saving)")
    loose = [d for d in disks if d.state.lower() == "unattached"]
    loose_spend = [d.price for d in loose if d.price is not None]
    item("Unattached disks", len(loose), "int")
    item("Unattached capacity, GB", sum(d.size_gb for d in loose), "int")
    item("Unattached spend, $/mo", r2(sum(loose_spend)) if loose_spend or not loose else None, "usd")
    if len(loose) > len(loose_spend):
        item("Unattached disks not priced", f"{len(loose) - len(loose_spend)} (see the Unattached disks sheet)")

    section("Next step")
    item("Please e-mail this file to", SEND_TO)
    return rows, bold


def about_rows(ran_in: int, guest_answer: str) -> list:
    reads = ("Azure Resource Graph (VMs, disks and resource groups), Azure Compute (VM size limits), Azure Monitor "
             "(disk metrics) and the public price list")
    rows = [
        ("What this is", "An estimate, not a quote. The figures come from your Azure and from public list prices "
                         "at the time of the run."),
        ("The saving", "For each Linux VM: used space is read inside the VM (df). Zoorik ElasticVol moves the "
                       "VM's data mounts onto one btrfs pool sized to the data: pool size "
                       f"= used + 20% (at least {HEADROOM_MIN_GIB} GB), rounded up, laid out as ElasticVol plans "
                       "it (2 anchor disks and 3 to 6 equal elastic disks; on VMs with few free data-disk slots, "
                       "1 anchor and up to 3 elastics). A Premium SSD v2 pool adds one 8 GB disk. In a Premium SSD "
                       f"pool every disk is a Premium SSD size and elastics are at least {ELASTIC_FLOOR_V1_GIB} GB. "
                       "Saving = today's list price of the disks the pool replaces - the pool's list price. After "
                       "the move, autopilot keeps the pool 75-85% full as the data grows and shrinks."),
        ("Pool disks", "Premium SSD v2 wherever the VM's region and zone offer it (as Azure lists it for your "
                       "subscription; a VM without a zone also needs a region where Microsoft supports nonzonal "
                       "Premium SSD v2), otherwise Premium SSD. Pool disks are locally redundant (LRS)."),
        ("Performance", "Measured peak = the highest hourly figure of the last 14 days in Azure Monitor, where each "
                        "hour's figure is its busiest minute (Maximum) of Composite Disk Read/Write Operations/sec "
                        "and Bytes/sec, reads and writes added and summed over the disks the pool replaces. Each "
                        "Premium SSD v2 disk includes 3,000 IOPS and 125 MB/s; more is priced in only when the "
                        "measured peak x 1.3, shared across the pool's disks, needs it, and never more than the VM "
                        "can drive. A Premium SSD disk has the fixed IOPS and MB/s of its size (bursting not "
                        "counted); a VM whose Premium SSD pool cannot carry the measured peak x 1.3 is not "
                        "estimated."),
        ("Prices", "Public Azure Retail Prices API (prices.azure.com): pay-as-you-go list prices in USD for each "
                   f"disk's region, read at run time. A month is {HOURS_PER_MONTH} hours. Premium SSD, Standard SSD "
                   "and Standard HDD are priced by tier (a Premium SSD set to a higher performance tier at that "
                   "tier); Premium SSD v2 and Ultra Disk by provisioned GB, IOPS and MB/s."),
        ("Not included", "Negotiated discounts, reservations, savings plans and credits; transaction charges, "
                         "bursting and snapshots; OS disks; unmanaged (VHD) disks; scale-set (uniform) instances; "
                         "subscriptions in other Azure directories."),
        ("Lanes", "Ready: ElasticVol can take the VM as it is. Needs work: a fix first (kernel 5.14 or later, "
                  "Azure Disk Encryption, free data-disk slots). Not supported: Windows; VMs in a resource group "
                  "another Azure service manages (such as Azure Databricks or an AKS node resource group); VM sizes "
                  f"with fewer than {MIN_VM_SLOTS} data-disk slots; RHEL-family Linux; no btrfs in the kernel; no "
                  "mounted data disk other than shared disks, which stay as they are. Not checked: disk usage was "
                  "not read."),
        ("Used space", "Used GB is the space used on the data-disk mounts the pool would take, from df inside the "
                       "VM. Mounted-disk utilization % = Used GB / the size of the data disks under those mounts; "
                       "Provisioned GB counts every data disk."),
        ("Blank cells", "A blank figure could not be known; the Reason or Note column says why."),
        ("Units", "GB means GiB (2^30 bytes), the unit Azure sizes disks in. MB/s means 10^6 bytes per second."),
    ]
    if ran_in:
        rows.append(("What ran", f"Read calls: {reads}; plus Azure Run Command in {ran_in} VM(s), described under "
                                 "Inside the VMs. Nothing else in your Azure was changed, and nothing was sent "
                                 "anywhere."))
        rows.append(("Inside the VMs", "With your yes, this command, which only reads, ran in each of those VMs "
                                       "through Azure Run Command (RunShellScript). On its side, Azure added its Run "
                                       "Command extension (RunCommandLinux) to VMs that did not have it, kept a copy "
                                       "of the command and its output under /var/lib/waagent on each VM, and logged "
                                       "each run in the activity log.\n" + GUEST_SCRIPT))
    else:
        rows.append(("What ran", f"Read calls only: {reads}. Nothing in your Azure was changed, and nothing was "
                                 "sent anywhere."))
        rows.append(("Inside the VMs", f"Nothing ran inside any VM ({guest_answer})."))
    rows.append(("Prepared by", f"Zoorik ElasticVol disk assessment {VERSION}. Run it again: {ONE_LINER}  "
                                f"Questions: {SEND_TO}"))
    return [list(r) for r in rows]


def build_sheets(run: dict, vms: List[Vm], disks: List[Disk]) -> list:
    names = {s.id: s.name for s in run["scanned"]}
    summary, bold = summary_rows(run, vms, disks)
    vm_list, vm_total = vm_rows(vms, names)
    loose, loose_total = unattached_rows(disks, names, run["now"])
    return [
        Sheet("Summary", [("Item", "text"), ("Value", "text")], summary, bold_rows=bold, widths=[44, 90],
              autofilter=False, preamble=[("Zoorik ElasticVol disk assessment", "title"), ("", "")]),
        Sheet("VMs", VM_COLS, vm_list, total=vm_total),
        Sheet("Disks", DISK_COLS, disk_rows(vms)),
        Sheet("Unattached disks", UNATTACHED_COLS, loose, total=loose_total),
        Sheet("About", [("Topic", "text"), ("Detail", "text")], about_rows(run["ran_in"], run["guest_answer"]),
              widths=[18, 120], autofilter=False),
    ]


# ---------------------------------------------------------------- xlsx writer

_FORMATS = {"text": 0, "int": 3, "usd": 164, "pct": 165}


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
    bold_rows: set = field(default_factory=set)


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
    text = escape(_BAD_XML.sub("", str(value))[:32767])
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def shown_length(value, fmt: str) -> int:
    if isinstance(value, tuple):
        value, fmt = value
    if value is None:
        return 0
    if isinstance(value, str):
        return len(value)
    if fmt == "int":
        return len(f"{value:,.0f}")
    if fmt == "pct":
        return len(f"{value * 100:.1f}%")
    if fmt == "usd":
        return len(f"${value:,.2f}")
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
    body = [(row, k in s.bold_rows) for k, row in enumerate(s.rows)] + ([(s.total, True)] if s.total else [])
    for n, (row, bold) in enumerate(body, head + 1):
        cells = []
        for i, (value, (_, fmt)) in enumerate(zip(row, s.columns)):
            if isinstance(value, tuple):
                value, fmt = value
            last = i == len(s.columns) - 1
            style = STYLE["wrap"] if s.widths and fmt == "text" and last and not bold else STYLE[(fmt, bold)]
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
            f'<styleSheet {_NS}><numFmts count="2"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/>'
            '<numFmt numFmtId="165" formatCode="0.0%"/></numFmts>'
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

def in_cloud_shell() -> bool:
    return bool(os.environ.get("ACC_CLOUD")) or os.environ.get("AZUREPS_HOST_ENVIRONMENT", "").startswith("cloud-shell")


LISTED_NAMES = 20


def prompt_text(candidates: List[Vm], names: Dict[str, str]) -> str:
    count = len(candidates)
    minutes = max(1, math.ceil(math.ceil(count / WORKERS) * RUN_COMMAND_SECONDS / 60))
    script = "\n".join("    " + line for line in GUEST_SCRIPT.splitlines())
    groups: Dict[str, List[str]] = {}
    for v in candidates:
        groups.setdefault(names.get(v.subscription, v.subscription), []).append(v.name)
    listed, left, rest, rest_subs = [], LISTED_NAMES, 0, 0
    for sub in sorted(groups):
        vm_names = sorted(groups[sub])
        if left <= 0:
            rest, rest_subs = rest + len(vm_names), rest_subs + 1
            continue
        shown = vm_names[:left]
        left -= len(shown)
        more = len(vm_names) - len(shown)
        listed.append(f"    {sub}: {', '.join(shown)}" + (f" and {more} more" if more else ""))
    if rest:
        listed.append(f"    and {rest} more in {rest_subs} other subscription(s)")
    idle = ("  Cloud Shell closes a session after 20 minutes without a key press. Press Enter every few\n"
            "  minutes until the report is written.\n\n") if in_cloud_shell() and minutes >= 10 else ""
    return (f"\nDisk usage\n"
            f"  {count} running Linux VM(s) have data disks:\n" + "\n".join(listed) + "\n\n"
            f"  Their saving needs the used space on those disks, which only the VM itself can report.\n"
            f"  With your yes, this command, which only reads, runs in each of them through Azure Run\n"
            f"  Command, {WORKERS} VMs at a time (about {minutes} min):\n\n"
            f"{script}\n\n"
            f"  On its side, Azure adds its Run Command extension (RunCommandLinux) to a VM that does not\n"
            f"  have it yet, keeps a copy of the command and its output under /var/lib/waagent on the VM,\n"
            f"  and records each run in the activity log. It needs the permission\n"
            f"  {RUN_COMMAND_PERMISSION} (Virtual Machine Contributor has it);\n"
            f"  a VM without it shows no saving, with the reason.\n\n"
            f"{idle}"
            f"Read disk usage inside these {count} VM(s) with Azure Run Command? [y/N] ")


def collect_facts(vms: List[Vm], disks: List[Disk]) -> Dict[str, Optional[Prices]]:
    """VM size limits per subscription and region, and list prices per region, read in parallel."""
    keys = sorted({(v.subscription, v.region) for v in vms if re.fullmatch(r"[a-z0-9]+", v.region)})
    regions = sorted({k[1] for k in keys} | {d.region for d in disks if re.fullmatch(r"[a-z0-9]+", d.region)})
    if not regions:
        return {}
    facts = dict(zip(keys, parallel(region_facts, keys, "Reading VM size limits, subscription-regions")))
    fetched = parallel(region_prices, regions, "Reading list prices, regions")
    prices = {r: p if isinstance(p, Prices) else None for r, p in zip(regions, fetched)}
    errors = sorted({str(p) for p in fetched if not isinstance(p, Prices)})
    if errors:
        say(f"List prices: {errors[0]}")
    for v in vms:
        f = facts.get((v.subscription, v.region))
        if isinstance(f, RegionFacts):
            limits = f.sizes.get(v.size.lower())
            if limits:
                v.slots, v.uncached_iops, v.uncached_mbps = limits
            else:
                v.facts_note = f"{v.size} not listed in {v.region}"
            v.pool_v2 = v2_offered(v, f)
        else:
            v.facts_note = str(f) if f else "region unknown"
    for d in disks:
        price_disk(d, prices)
    return prices


def unused_path(folder: str, stem: str) -> str:
    """A report path in the folder that no earlier report holds, so nothing is overwritten."""
    path, n = os.path.join(folder, stem + ".xlsx"), 1
    while os.path.exists(path):
        n += 1
        path = os.path.join(folder, f"{stem}-{n}.xlsx")
    return path


def assessment() -> int:
    say(f"Zoorik ElasticVol disk assessment {VERSION}")
    if os.name == "nt":
        say("On Windows, run this in Azure Cloud Shell (shell.azure.com) or in WSL.")
        return 1
    RUN_COMMAND_STARTED.clear()
    try:
        directory, subs, skipped = signed_in()
    except Failed as e:
        if "not installed" in str(e):
            say("This needs the Azure CLI (az). Open Azure Cloud Shell (shell.azure.com) and paste the line there.")
        else:
            say(f"Azure: {e}, then paste the line again.")
        return 1
    if not subs:
        say("No enabled Azure subscription is visible to this sign-in. Run az login, then paste the line again.")
        return 1
    say(f"Directory: {directory}")
    for s in skipped:
        say(f"Not scanned: {s.name} ({s.id}), {s.note}")
    vms, disks = inventory(subs)
    empty = 0
    for s in subs:
        if s.note:
            say(f"Subscription {s.name}: not scanned, {s.note}")
        elif s.vms or s.disks:
            say(f"Subscription {s.name}: {s.vms} VMs, {s.disks} disks")
        else:
            empty += 1
    if empty:
        say(f"{empty} subscription(s) with no VMs or disks.")
    if not vms and not disks:
        if all(s.note for s in subs):
            say("No subscription could be read, so there is nothing to report. Fix the reason above, "
                "then paste the line again.")
            return 1
        say("No VMs or managed disks in the subscriptions scanned, so there is nothing to report.")
        return 0
    prices = collect_facts(vms, disks)

    candidates = []
    for v in vms:
        if not v.data or not v.linux or screen(v):
            continue
        if not v.running:
            v.guest_note = "VM is not running"
        elif not any(dd.disk for dd in v.data):
            v.guest_note = "no managed data disk"
        else:
            candidates.append(v)
    answer, guest_answer = None, "no running Linux VM with data disks to read"
    if candidates:
        answer = ask(prompt_text(candidates, {s.id: s.name for s in subs}))
        if answer is None:
            say("\nNo terminal to ask on, so disk usage inside VMs is not read.")
        guest_answer = {True: "", False: "declined at the prompt", None: "no terminal to ask on"}[answer]
    if answer:
        for v, result in zip(candidates, parallel(read_guest, candidates, "Reading disk usage, VMs")):
            if isinstance(result, Guest):
                v.guest = result
            else:
                v.guest_note = str(result)
    else:
        for v in candidates:
            v.guest_note = guest_answer

    measured = [dd.disk for v in vms if v.linux for dd in v.data if dd.disk]
    for d, result in zip(measured, parallel(disk_metrics, measured, f"Reading {METRIC_DAYS} days of disk metrics")):
        if isinstance(result, dict) and result:
            d.hourly = result
        else:
            d.metric_note = f"metrics not readable: {result}" if isinstance(result, Failed) else \
                f"no metric data in the last {METRIC_DAYS} days"
    for v in vms:
        if not v.linux:
            for dd in v.data:
                if dd.disk:
                    dd.disk.metric_note = "not measured (Windows VM)"
        try:
            assess(v, prices)
        except Exception as e:  # one VM with unexpected data must not stop the report
            v.lane, v.reasons = "", [f"not assessed: unexpected data ({type(e).__name__})"]

    now = datetime.now(timezone.utc)
    run = {"directory": directory, "scanned": subs, "skipped": skipped, "now": now,
           "when": now.strftime("%Y-%m-%d %H:%M UTC"), "ran_in": len(RUN_COMMAND_STARTED),
           "guest_answer": guest_answer}
    sheets = build_sheets(run, vms, disks)
    stem = f"zoorik-disk-assessment-{now:%Y%m%d-%H%M}"
    path = ""
    for folder in (os.getcwd(), os.path.expanduser("~")):
        try:
            path = unused_path(folder, stem)
            write_xlsx(path, sheets)
            break
        except OSError:
            path = ""
    if not path:
        say(f"Could not write {stem}.xlsx here or in your home folder.")
        return 1

    with_data = [v for v in vms if v.data]
    attached = attached_once(with_data)
    spend = priced(attached)
    unpriced = len(attached) - len(spend)
    savings = [v.saving for v in with_data if v.saving]
    loose = [d for d in disks if d.state.lower() == "unattached"]
    loose_spend = [d.price for d in loose if d.price is not None]
    total_spend = r2(sum(spend)) if spend else None
    say()
    say(f"VMs:               {len(vms)} scanned, {len(with_data)} with data disks")
    say(f"Data-disk spend:   {money(total_spend)}" + (" per month at list prices" if spend else "")
        + (f" ({unpriced} data disk(s) not priced)" if unpriced else ""))
    if savings:
        share = f" ({sum(savings) / total_spend:.0%} of data-disk spend)" if total_spend else ""
        say(f"Estimated saving:  {money(r2(sum(savings)))} per month on {len(savings)} VM(s){share}")
    else:
        say(f"Estimated saving:  {no_saving_text(vms)}")
    say(f"Unattached disks:  {len(loose)}" + (f", {money(r2(sum(loose_spend)))} per month" if loose_spend else "")
        + (f" ({len(loose) - len(loose_spend)} not priced)" if len(loose) > len(loose_spend) else ""))
    say()
    say(f"Report: {path}")
    if in_cloud_shell():
        say(f"Download it: Cloud Shell toolbar > Manage files > Download, then enter {path}")
    say(f"Please e-mail it to {SEND_TO}")
    return 0


def stop_note() -> str:
    """What the run had changed in Azure when it stopped."""
    if not RUN_COMMAND_STARTED:
        return "Nothing in your Azure was changed."
    return (f"Azure Run Command had already started in {len(RUN_COMMAND_STARTED)} VM(s) (see the activity log); "
            "nothing else in your Azure was changed.")


def main() -> int:
    try:
        return assessment()
    except KeyboardInterrupt:
        say(f"\nStopped. {stop_note()}")
        os._exit(130)  # leave at once: never wait for calls still in flight
    except Failed as e:
        say(f"\nStopped: {e}. {stop_note()}")
        return 1
    except Exception as e:  # a prospect sees one plain line, never a traceback
        say(f"\nStopped by an unexpected error ({type(e).__name__}). {stop_note()} "
            f"Please send this line to {SEND_TO}.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
