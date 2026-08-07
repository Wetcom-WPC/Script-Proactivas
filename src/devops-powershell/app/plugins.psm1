$folder = $global:CONFIG.PLUGINS_FOLDER

function Start-Plugins {
    Get-ChildItem -Path $global:CONFIG.PLUGINS_FOLDER | 
        Where-Object{$_.Extension -eq ".psm1"} |
        ForEach-Object {
            Import-Module -Global "$_"
            $modules = Get-Command -Module $_.BaseName | Get-Help | 
                Select-Object Name, SYNOPSIS, COMPONENT, ROLE | 
                Where-Object {$_.role -eq 'ui'}
            
            foreach ($m in $modules) {
                Add-Member -InputObject $m -MemberType NoteProperty -Name 'checked' -Value $false
                $global:PLUGINS_MODULES += $m
            }    
    }
}

