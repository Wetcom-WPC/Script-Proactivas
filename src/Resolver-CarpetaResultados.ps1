# Calcula y crea la carpeta de resultados para esta corrida: resultados/<yyyy-MM-dd>,
# o resultados/<yyyy-MM-dd> (1), (2), (3)... si ya existe una del mismo dia.
# Imprime solo el nombre de la carpeta (sin ruta) para que lanzador.bat lo capture.
param(
    [Parameter(Mandatory = $true)]
    [string]$BaseDir
)

$fecha = Get-Date -Format "yyyy-MM-dd"
$directorioResultados = Join-Path -Path $BaseDir -ChildPath "resultados"

$nombreCarpeta = $fecha
$candidato = Join-Path -Path $directorioResultados -ChildPath $nombreCarpeta

if (Test-Path $candidato) {
    $n = 1
    do {
        $nombreCarpeta = "$fecha ($n)"
        $candidato = Join-Path -Path $directorioResultados -ChildPath $nombreCarpeta
        $n++
    } while (Test-Path $candidato)
}

New-Item -ItemType Directory -Path $candidato -Force | Out-Null

Write-Output $nombreCarpeta
