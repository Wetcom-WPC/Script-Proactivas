$global:UI_LINEH = "═"
$global:UI_UPLEFT = "╔"
$global:UI_UPRIGHT = "╗"
$global:UI_LINEV = "║"
$global:UI_DOWNLEFT = "╚"
$global:UI_DOWNRIGHT = "╝"
$global:UI_LINEVLEFTCROSS = "╠"
$global:UI_LINEVRIGHTROSS = "╣"

$global:ExecuteEvent = @()
$global:ClearConnectionsEvent = @()


function Register-ExecuteEventSubscription($commandName) {
    $global:ExecuteEvent += $commandName
}

function Invoke-ExecuteEvent {
    foreach ($subscription in $global:ExecuteEvent) {
        &$subscription
    }
}

function Register-ClearConnectionsEventSubscription($commandName) {
    $global:ClearConnectionsEvent += $commandName
}

function Invoke-ClearConnectionsEvent {
    foreach ($subscription in $global:ClearConnectionsEvent) {
        &$subscription
    }
}

function Start-UI {
    Show-Menu -Plugins $global:PLUGINS_MODULES

    while($true) {
        Start-Sleep -MilliSeconds 200
        if ($Host.UI.RawUI.KeyAvailable) {
            $keydown = $Host.UI.RawUI.ReadKey("IncludeKeyUp")
            $command = $keydown.Character
            if($keydown.VirtualKeyCode -ne 13){
                Write-Host -NoNewline "`b `b"
            }
        } 
        if ($command -eq "x") { 
            exit 1
        } elseif ($command -Match "[1-9]") {
            $index = [convert]::ToInt32($command, 10)
            $index--
            $global:PLUGINS_MODULES[$index].checked = !$global:PLUGINS_MODULES[$index].checked
            Show-Menu -Plugins $global:PLUGINS_MODULES
        } elseif ($command -eq "e"){
            Invoke-ExecuteEvent
            Show-Menu -Plugins $global:PLUGINS_MODULES
            exit 0
        } elseif ($command -eq "c"){
            Invoke-ClearConnectionsEvent
            $global:PLUGINS_MODULES | ForEach-Object {$_.checked = $false}
            Show-Menu -Plugins $global:PLUGINS_MODULES
        } 
        
        # Resize
        if ($global:ScreenSize -ne $host.UI.RawUI.BufferSize.Width) {
            $global:ScreenSize = $host.UI.RawUI.BufferSize.Width
            Show-Menu -Plugins $global:PLUGINS_MODULES
        }

        $command = $null
    }
}

function Show-Menu($Plugins) 
{
    Clear-Host 
    $index = 1
    $width = $host.UI.RawUI.BufferSize.Width
    
    $lineColor = 'Green'

    $lineaSuperior = $global:UI_UPLEFT + ($global:UI_LINEH * ($width - 2)) + $global:UI_UPRIGHT
    Write-Host $lineaSuperior -ForegroundColor $lineColor

    $textoEncabezado = $global:CONFIG.APP_HEADER + " " + $global:APP_VERSION
    $paddedText = (" " + $textoEncabezado + (" " * $width)).Substring(0, $width - 2)
    Write-Host $global:UI_LINEV -NoNewline -ForegroundColor $lineColor
    Write-Host $paddedText -NoNewline -ForegroundColor 'Green' # Color del texto
    Write-Host $global:UI_LINEV -ForegroundColor $lineColor
    
    $lineaSeparadora = $global:UI_LINEVLEFTCROSS + ($global:UI_LINEH * ($width - 2)) + $global:UI_LINEVRIGHTROSS
    Write-Host $lineaSeparadora -ForegroundColor $lineColor

    foreach ($conn in $global:connections) {
        $textoConexion = "$($conn.component): $($conn.host)"
        $paddedText = (" " + $textoConexion + (" " * $width)).Substring(0, $width - 2)
        Write-Host $global:UI_LINEV -NoNewline -ForegroundColor $lineColor
        Write-Host $paddedText -NoNewline -ForegroundColor 'Green' 
        Write-Host $global:UI_LINEV -ForegroundColor $lineColor
    }

    Write-Host $lineaSeparadora -ForegroundColor $lineColor

    foreach ($plugin in $Plugins) {
        $text = $plugin.Synopsis
        $textColor = if ($plugin.checked) {'Red'} else {'Green'}
        $checked = if ($plugin.checked) {"X"} else {" "}
        $optionText = "[$checked] ($index) $text"
        $index++
        
        $paddedText = (" " + $optionText + (" " * $width)).Substring(0, $width - 2)
        Write-Host $global:UI_LINEV -NoNewline -ForegroundColor $lineColor
        Write-Host $paddedText -NoNewline -ForegroundColor $textColor 
        Write-Host $global:UI_LINEV -ForegroundColor $lineColor
    }

    Write-Host $lineaSeparadora -ForegroundColor $lineColor
    
    $textoComandos = "(c) Clear data  |  (e) Execute |  (x) Exit"
    $paddedText = (" " + $textoComandos + (" " * $width)).Substring(0, $width - 2)
    Write-Host $global:UI_LINEV -NoNewline -ForegroundColor $lineColor
    Write-Host $paddedText -NoNewline -ForegroundColor 'Green' 
    Write-Host $global:UI_LINEV -ForegroundColor $lineColor

    $lineaInferior = $global:UI_DOWNLEFT + ($global:UI_LINEH * ($width - 2)) + $global:UI_DOWNRIGHT
    Write-Host $lineaInferior -ForegroundColor $lineColor
}


function Write-Title($text) {
	Write-Host " "
    "=" * $text.Length
    $text
    "=" * $text.Length
}

function Show-Progress($total, $index){
	$current = [math]::Round(($index / $total) * 100)
	$whitespace = 3 - $current.ToString().Length
	Write-Host (" " * $whitespace  + $current + "%") -NoNewline
    
    if ($current -lt 100) {
        Write-Host "`b`b`b`b" -NoNewline
    } else {
        Write-Host ""
    }
}