Write-Host "Actualizando ESXi vNic references..."
Invoke-WebRequest -Uri "http://www.virten.net/repo/vmware-iohcl.json" -OutFile "./data/vmware-iohcl.json"

Write-Host "Actualizando ESXi Releases..."
Invoke-WebRequest -Uri "http://www.virten.net/repo/esxiReleases.json" -OutFile "./data/esxiReleases.json"

Write-Host "Actualizando VMWare HCL..."
Invoke-WebRequest -Uri "https://www.virten.net/repo/vmware-hcl.json" -OutFile "./data/vmware-hcl.json"

Write-Host "Actualizando vSAN VMWare HCL..."
Invoke-WebRequest -Uri "http://partnerweb.vmware.com/service/vsan/all.json" -OutFile "./data/vsan-hcl.json"

Write-Host "Actualizando ESXi required VM tools references..."
$content = (Invoke-WebRequest -Uri  https://packages.vmware.com/tools/versions).content -split "`n"
$content = $content | Where-Object {$_ -notmatch "^#"}

$toolsRef= @()
FOReach($line in $content)
{
    $data = $line.split()| Where-Object {$_}
    if ($null -eq $data -or $data.length -ne 4) {continue}
    $toolsRef += [PSCustomObject] @{
        toolsBuild = $data[0]
        esxiVersion = $data[1]
        toolsVersion = $data[2]
        esxiBuild = $data[3]
    }
}
$toolsRef | ConvertTo-Json | Out-File "./data/vmtools-ref.json"

Write-Host "Done"
