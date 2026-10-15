# Environments the proactiva covers only in part.
#
# Azure VMware Solution (AVS): Microsoft runs the hosts, their certificates, the appliance
# accounts, the licenses and most of the vCenter settings, and the customer's user has no rights
# over them, so what the checks would find there is not something the customer can act on. What
# runs inside the virtual machines is the customer's: for AVS vCenters only the VMs and their
# snapshots are collected, and only the VM-level sheets reach the checks (mixed with the VMs of
# the other vCenters). If the data of a past run has AVS rows in other sheets, the report leaves
# them out; the raw Excel keeps everything.

# AVS vCenters are named "vc.<id>.<region>.avs.azure.com"
function Test-IsAvsVCenter {
    param([string] $Name)

    return ([string]$Name -match '\.avs\.azure\.com\.?$')
}

# Raw Excel sheets whose AVS rows are analyzed
function Get-AvsSheets {
    return @('VM', 'vNetwork', 'Snapshot')
}

Export-ModuleMember -Function Test-IsAvsVCenter, Get-AvsSheets
