# Environments the proactiva does not cover.
#
# Azure VMware Solution (AVS): Microsoft runs the hosts, their certificates, the appliance
# accounts and most of the vCenter settings, and the customer's user has no rights over them,
# so what the checks would find there is not something the customer can act on. AVS vCenters
# are not collected, and if they are present in the data of a past run they are left out of the
# annex and the checklist (the raw Excel keeps them).

# AVS vCenters are named "vc.<id>.<region>.avs.azure.com"
function Test-IsAvsVCenter {
    param([string] $Name)

    return ([string]$Name -match '\.avs\.azure\.com\.?$')
}

Export-ModuleMember -Function Test-IsAvsVCenter
