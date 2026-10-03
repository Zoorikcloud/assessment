# Zoorik assessments

Zoorik's free assessments for Coral (Kubernetes) and Amoeba (Azure block storage). You paste one line. The script writes one Excel file that estimates what you could save, and you e-mail that file to founders@zoorik.com.

- [Kubernetes: `k8s.py`](#kubernetes-k8spy) covers AKS, EKS, or any cluster kubectl can reach.
- [Azure disks: `disks.py`](#azure-disks-diskspy) covers the data disks of your Azure Linux VMs.
- [Where the figures come from](#where-the-figures-come-from)

Each script is one Python file. It needs Python 3.8 or later and uses only the standard library. Azure Cloud Shell and AWS CloudShell already have Python 3 and the CLIs the scripts call. Each `zoorik.com/assess/...` URL redirects to the file of the same name in this repo, so the code you read here is the code that runs.

---

## Kubernetes: k8s.py

```bash
curl -fsSL https://zoorik.com/assess/k8s | python3 -
```

### Where to run it

- **Azure Cloud Shell (Bash):** it assesses every AKS cluster in the current subscription.
- **AWS CloudShell:** it assesses every EKS cluster in every enabled region.
- **Any other shell:**
  - it assesses the AKS clusters in the current subscription if `az` is signed in;
  - it assesses the EKS clusters if `aws` is signed in;
  - if neither finds a cluster, it assesses the current kubectl context.

  Outside a Cloud Shell it needs `python3`, `kubectl`, and the `az` or `aws` CLI for that cloud.

A run takes a little over 3 minutes. It reads each cluster, then samples usage for 3 minutes. Each private AKS cluster adds about a minute, and many or large clusters take longer.

Both Cloud Shells close a session after about 20 minutes without a key press, even while the script runs. On a long run, press Enter now and then until the report path appears. The script warns you when a run may be long.

### What it reads

- **In each cluster, with `kubectl get`:** nodes, pods, ReplicaSets, Jobs, HPAs, PDBs and PersistentVolumeClaims.
- **With `kubectl top`:** node and container usage, every 30 seconds. It keeps the peak.
- **What it keeps:**
  - names, owners, requests, limits and volume claim names;
  - labels, used only for matching;
  - each node's machine type, zone, size, and whether it is Spot or on-demand.

  It never reads Secrets or ConfigMaps. It never keeps env values, annotations or pod IPs.
- **With `az` or `aws`:** the signed-in account, its clusters, node pools and regions. It also writes each cluster's credentials to a temporary kubeconfig.
- **Prices:**
  - for AKS, the public Azure Retail Prices API (prices.azure.com);
  - for EKS, the AWS Price List API and EC2 Spot price history, through `aws`.

### Beyond reads: one short-lived pod, only in AKS clusters it can't read directly

The script first tries kubectl. Some AKS clusters can't be read that way from where you run it:

- a private AKS cluster, from outside its network;
- an AKS cluster that kubectl can't reach from there;
- an AKS cluster kubectl can't sign in to, because kubelogin is not installed.

For these, the script uses `az aks command invoke`. Azure starts a short-lived pod of its own in the cluster's `aks-command` namespace, and the same kubectl reads run in that pod.

- Azure documents this pod as requesting 200m CPU and 500 MiB of memory. The report leaves it out.
- Azure records each call in the activity log.
- Azure caps each call's output at 512 KB.
- A cluster read this way gives one usage snapshot instead of 3 minutes of samples.

`--no-invoke` turns this off. kubectl then has to reach those clusters directly, and any it can't reach are listed as not assessed, with the reason.

Apart from this pod, the script only reads, and it writes only the files below.

### What it sends

Nothing. There is no upload and no telemetry, and you send the file yourself. It switches off the Azure CLI's own usage telemetry for its calls, without changing your az settings. Its only calls go to:

- your clusters;
- your cloud's own APIs;
- the public Azure price list, which receives a region and a machine type for each lookup.

### What it writes

- **One Excel file,** `zoorik-k8s-assessment-YYYYMMDD-HHMM.xlsx`.
  - In a Cloud Shell it goes in your home folder. Elsewhere it goes in the current directory, or in your home folder if the current directory can't be written. `--out` chooses another place.
  - Its sheets are Summary, Node pools, Nodes, Workloads and Assumptions.
- **Working files:** kubectl's cache and the temporary kubeconfig. They go in a temporary folder that is deleted when the script exits. Your own kubeconfig is left as it is.

### Access it needs

**Azure:**
- read access to the subscription's AKS clusters;
- the *Azure Kubernetes Service Cluster User Role* on each cluster;
- for `command invoke`, also `Microsoft.ContainerService/managedClusters/runcommand/action` and `Microsoft.ContainerService/managedClusters/commandResults/read`.

**AWS:**
- `eks:ListClusters`, `eks:DescribeCluster`, `eks:ListNodegroups` and `eks:DescribeNodegroup`;
- `ec2:DescribeRegions` and `ec2:DescribeSpotPriceHistory`;
- `pricing:GetProducts`;
- an EKS access entry for the identity.

The identity that created the cluster usually has an access entry already. For any other identity, choose one of these:
- Associate `AmazonEKSAdminViewPolicy` at cluster scope. This policy can also read Secrets, but the script never asks for them.
- Put a Kubernetes group in the access entry, and bind that group to the role below, which allows only `get` and `list`.

`AmazonEKSViewPolicy` is not enough: it can't list nodes or read metrics.

**Inside each cluster:** the identity needs `get` and `list`, cluster-wide, on the resources the script reads. If it doesn't have them, a cluster admin can create this role and bind it. Replace `GROUP` with the Kubernetes group from the EKS access entry, or with the Entra ID group's object ID on AKS:

```bash
kubectl create clusterrole zoorik-assess-read --verb=get,list --resource=nodes,pods,persistentvolumeclaims,replicasets.apps,jobs.batch,horizontalpodautoscalers.autoscaling,poddisruptionbudgets.policy,nodes.metrics.k8s.io,pods.metrics.k8s.io
kubectl create clusterrolebinding zoorik-assess-read --clusterrole=zoorik-assess-read --group=GROUP
```

When the script can't read a cluster, the run goes on to the next one, and the Status column says why.

### Options

Put options after the dash:

```bash
curl -fsSL https://zoorik.com/assess/k8s | python3 - --anonymize
```

| Option | Effect |
|---|---|
| `--anonymize` | Replace namespace, workload and node names with short hashes, salted for this run. Node pool names stay. |
| `--anonymize-clusters` | Also replace cluster names. |
| `--sample-minutes N` | Sample usage for N minutes (default 3; 0 takes one sample). |
| `--subscription ID` | Search this Azure subscription instead of the current one. |
| `--region NAME` | Search this AWS region; repeat it for more (default: every enabled region). |
| `--discover azure`, `--discover aws` | Find clusters through `az` or `aws`. |
| `--context NAME` | Assess this kube context; repeat it for more. |
| `--all-contexts` | Assess every context in your kubeconfig. A cluster reached through several contexts is counted once. |
| `--no-invoke` | Never use `az aks command invoke`. |
| `--no-prices` | Skip the price lookups, so there are no cost or savings figures. |
| `--out PATH` | Write the Excel file to this file or folder. A path ending in `/` is a folder, created if needed; a file name gets `.xlsx` added if it lacks it. |

`--subscription`, `--region`, `--discover`, `--context` and `--all-contexts` replace the automatic choice of clusters: only what you name is assessed. `--help` lists every option.

### Download the file and send it

When the script finishes, it prints the file's full path.

- **Azure Cloud Shell:** on the toolbar, choose **Manage files > Download**, then enter that path.
- **AWS CloudShell:** choose **Actions > Download file**, then enter that path. This works only from the full CloudShell page. If **Actions** is missing, open CloudShell in a new browser tab.

Then e-mail the file to **founders@zoorik.com**.

### How the estimate works

**1. Right-size the requests.** For each container:
- the recommended request is the peak sampled usage × 1.3;
- it is at least 10m CPU and 32 MiB of memory;
- it is never above today's request, unless today's request is zero.

Exceptions:
- with no usage data, today's request stays;
- init containers keep their requests, except native sidecars (`restartPolicy: Always`), which are right-sized like the other running containers;
- pods with pod-level requests keep them as they are.

**2. Count the nodes needed.** This is worked out per node pool and machine type. DaemonSet and static pods count as per-node overhead. Nodes needed is the largest of these:
- ceil(recommended CPU / ((allocatable CPU per node − overhead per node) × 0.85));
- the same for memory;
- ceil(workload pods / ((allocatable pods per node − DaemonSet and static pods per node) × 0.85)).

It is never more than today's count. Each pool keeps at least its autoscaler minimum, and at least 1 node, across its machine types. Where the cloud API gives no minimum for a pool (a cluster read through a kube context, or a node group the API doesn't list), the floor is 1 node and the Status column says so.

**3. Price the difference.**

Savings per month = (nodes today − nodes needed) × the node's hourly list price × 730 hours.

**Prices** are USD list prices:
- **Azure:** Linux pay-as-you-go, or the Spot meter for Spot nodes;
- **AWS:** Linux on-demand with shared tenancy, or for Spot nodes the current Spot price averaged across zones.

They leave out discounts, reservations, savings plans and credits. Nodes whose labels don't say Spot or on-demand, such as self-managed EKS node groups, are not priced.

**Usage:**
- Three minutes of samples can miss peaks. Use `--sample-minutes` for a longer window.
- Without metrics-server there is no usage. Requests then stay as they are, and only surplus nodes count.

**Not included:**
- Further savings: moving to Spot, moving to cheaper or newer machine types, and packing across pools.
- Other costs: storage, network, load balancers, control-plane fees and licences. Windows nodes are not priced.
- Pending pods.

**Limits:**
- **Private EKS:** a private EKS API can't be reached from CloudShell. Run the line from a machine inside the VPC instead, such as an EC2 instance. A CloudShell VPC environment won't do: it can't download files. Run from CloudShell, the report lists that cluster's node groups and their cost from the AWS API, but has no savings estimate for it.
- **Large clusters through `command invoke`:** these are read one resource kind at a time. If one kind is still over 512 KB, the Status column names it.

---

## Azure disks: disks.py

```bash
curl -fsSL https://zoorik.com/assess/disks | python3 -
```

### Where to run it

Run it in Azure Cloud Shell (Bash), or in a Linux or macOS shell with `python3` and a signed-in `az`. On Windows, use Cloud Shell or WSL.

Cloud Shell closes a session after 20 minutes without a key press. On a long run, press Enter every few minutes until the report is written.

It assesses every enabled subscription in the Azure directory (tenant) you are signed in to. Subscriptions in other directories are listed as not scanned. To include one, run `az login --tenant <tenant-id>`, then run the line again.

### What it reads

- **Azure Resource Graph:** your VMs and managed disks, and which resource groups another Azure service manages.
- **Azure Compute:**
  - each VM size's data-disk slots;
  - each VM size's uncached disk IOPS and throughput limits;
  - which zones offer Premium SSD v2 to your subscription.
- **Azure Monitor:** 14 days of IOPS and throughput for each data disk of a Linux VM, as each hour's busiest minute. The metrics are Composite Disk Read/Write Operations/sec and Bytes/sec, with the Maximum aggregation.
- **prices.azure.com:** the public list prices of managed disks.

### Beyond reads: Azure Run Command, only if you answer y

The saving needs the used space on each data disk, and only the VM itself can report that.

The script first finds the running Linux VMs that have managed data disks and pass its first checks. If there are any, it lists them by subscription, prints the command below, says how long it will take, and asks once:

```
Read disk usage inside these N VM(s) with Azure Run Command? [y/N]
```

**If you answer y:**
- The command runs in each of those VMs through Azure Run Command (`RunShellScript`), 8 VMs at a time.
- It reads:
  - the kernel version and the Linux distribution;
  - whether btrfs is available;
  - which mounts sit on each data disk;
  - used space, from `df`.
- On its side, Azure:
  - adds its Run Command extension (RunCommandLinux) to any VM that doesn't have it yet;
  - keeps a copy of the command and its output under `/var/lib/waagent` on the VM;
  - records each run in the activity log.

**If you answer anything else, or there is no terminal to ask on:** nothing runs inside any VM. Those VMs show as Not checked, with no saving.

**Ctrl-C** stops the script at once, and no new Run Command starts after it. Run Commands already started finish in Azure, and the script's last line says how many VMs they reached.

This is the command:

```sh
echo S
echo "K $(uname -r)"
grep -E '^(ID|ID_LIKE|VERSION_ID)=' /etc/os-release
if grep -qw btrfs /proc/filesystems; then echo "B loaded"; elif modinfo -n btrfs >/dev/null 2>&1; then echo "B module"; else echo "B absent"; fi
for l in /dev/disk/azure/scsi1/lun*; do n=${l##*/lun}; case $n in ''|*[!0-9]*) continue;; esac; echo "L $n $(lsblk -nro MOUNTPOINT "$l" | tr '\n' ' ')"; done
timeout 30 df -P -B1 -T -l | awk '$1 ~ /^\/dev\// && $1 !~ /^\/dev\/loop/ {m=$7; for (i=8; i<=NF; i++) m=m" "$i; print "D", $2, $3, $4, m}'
echo E
```

Apart from this step, the script only reads, and it writes only the file below.

### What it sends

Nothing. There is no upload and no telemetry, and you send the file yourself. It switches off the Azure CLI's own usage telemetry for its calls, without changing your az settings. Its only calls go to:

- your Azure APIs;
- the public Azure price list, which receives a region name for each lookup.

### What it writes

One Excel file, `zoorik-disk-assessment-YYYYMMDD-HHMM.xlsx` (UTC).

- It goes in the current directory, or in your home folder if the current directory can't be written.
- It never overwrites an earlier report: a second run in the same minute adds `-2`, `-3` and so on.
- Its sheets are Summary, VMs, Disks, Unattached disks and About.
- It names your directory, subscriptions, resource groups, VMs, disks and mount points.

### Access it needs

- Reader on the subscriptions.
- For the Run Command step, `Microsoft.Compute/virtualMachines/runCommand/action`. Virtual Machine Contributor has it. A VM where this permission is missing shows no saving, with the reason.

### Options

None.

### Download the file and send it

In Azure Cloud Shell, on the toolbar, choose **Manage files > Download**, then enter the path the script prints. Then e-mail the file to **founders@zoorik.com**.

### How the estimate works

The script estimates a saving for each Linux VM whose disk usage was read.

1. **Used space** is the space used on the VM's data-disk mounts. Each mount is rounded up to a whole GB. A shared disk (one attached to several VMs) is never pooled; it stays as it is, and the totals count it once.
2. **Pool size** = used + 20% of used (at least 50 GB), rounded up.
3. **Disk type:** the pool goes on Premium SSD v2 where the VM's region and zone offer it, as Azure lists it for your subscription. A VM without a zone also needs a region where Microsoft supports nonzonal Premium SSD v2. Otherwise the pool goes on Premium SSD. Pool disks are locally redundant (LRS). Where today's disks are zone-redundant (ZRS), the VM's Reason says that part of the saving is that difference.
4. **Layout, as Amoeba plans it:**
   - The pool has 3 to 6 equal elastic disks. Pools of 512 GB and over also have 2 anchor disks.
   - On a VM with few free data-disk slots, the pool has 1 anchor and up to 3 elastics.
   - The pool is built beside today's disks, so it must fit the free slots, with one slot kept spare.
   - A Premium SSD v2 pool adds one 8 GB disk.
   - In a Premium SSD pool, each disk is a Premium SSD size, and elastics are at least 128 GB.
5. **Performance:**
   - The measured peak is the highest hourly figure of the last 14 days. Each hour's figure is its busiest minute (the Maximum), with reads and writes added and summed over the disks the pool replaces.
   - Each Premium SSD v2 disk includes 3,000 IOPS and 125 MB/s.
   - More is priced in only when the measured peak × 1.3, shared across the pool's disks, needs it. Even then it is never more than the VM itself can drive.
   - A Premium SSD pool has the fixed IOPS and MB/s of its disk sizes, without bursting. If those together can't carry the measured peak × 1.3 (or what the VM can drive, if less), the VM is not estimated.
   - Without metrics for the disks the pool replaces, the VM is not estimated.
6. **Saving per month** = today's list price of the disks the pool replaces − the pool's list price. Where the pool costs no less, the VM shows No saving.

**Prices:**
- **Source:** the public Azure Retail Prices API, using pay-as-you-go USD list prices for each disk's region, read at run time.
- **Premium SSD, Standard SSD and Standard HDD:** priced by billed tier. A Premium SSD set to a higher performance tier is priced at that tier.
- **Premium SSD v2 and Ultra Disk:** priced by provisioned GB, IOPS and MB/s, over a 730-hour month.

**Not included:**
- discounts, reservations, savings plans and credits;
- transaction charges, bursting and snapshots;
- OS disks;
- unmanaged (VHD) disks;
- scale-set (uniform) instances;
- subscriptions in other directories.

Data disks without a mounted file system stay as they are. Unattached disks are listed with their cost on their own sheet, outside the Amoeba saving.

**Lanes** (the Lane column on the VMs sheet):

| Lane | Meaning |
|---|---|
| Ready | Amoeba can take the VM as it is. |
| Needs work | The VM needs a fix first: a kernel below 5.14, Azure Disk Encryption, or not enough free data-disk slots. |
| Not supported | Windows; a VM in a resource group another Azure service manages, such as Azure Databricks or an AKS node resource group; fewer than 8 data-disk slots; RHEL-family Linux; no btrfs in the kernel; or no mounted data disk other than shared disks. |
| Not checked | Disk usage was not read. |

**Units:** GB means GiB (2^30 bytes), the unit Azure sizes disks in. MB/s means 10^6 bytes per second.

---

## Where the figures come from

Every figure in either file is one of these:
- read from your cloud or your clusters;
- read from the public price lists;
- computed from those by the method above.

When a figure can't be known, its cell is blank and the row says why. In the k8s.py report, see the Status and Price note columns; a blank limit there means at least one container has no limit. In the disks.py report, see the Reason and Note columns. Nothing is guessed or filled in.

The figures are estimates at list prices, not a quote.

Questions: founders@zoorik.com.
