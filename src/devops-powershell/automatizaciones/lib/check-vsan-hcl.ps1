

Function Get-VsanHclDatabase {
<#
	.SYNOPSIS
		This function will allow you to view and download the VSAN Hardware Compatability List (HCL) Database
	
	.DESCRIPTION
		Use this function to view or download the VSAN HCL
	.EXAMPLE
        View the latest online HCL Database from online source
		PS C:\> Get-VsanHclDatabase | Format-Table
	.EXAMPLE
        Download the latest HCL Database from online source and store locally
		PS C:\> Get-VsanHclDatabase -filepath ~/hcl.json
#>
Param (
    $vid,
    $did,
    $svid,
    $ssid,
    $vsanhcl
    )
Process {
    Foreach ($entry in $vsanhcl.data.controller) {
        If (($vid -eq $entry.vid) -and ($did -eq $entry.did) -and ($svid -eq $entry.svid) -and ($ssid -eq $entry.ssid) ) {
            $entry.vcglink
        }
    }
}
}