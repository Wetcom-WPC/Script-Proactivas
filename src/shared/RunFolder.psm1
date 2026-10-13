# Every run writes into its own folder: resultados\<yyyy-MM-dd>, or resultados\<yyyy-MM-dd> (1),
# (2)... when that day already has one (same numbering Windows uses for duplicate files).

function Get-ResultsRoot {
    return Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent | Split-Path -Parent) -ChildPath "resultados"
}

# Creates the folder for a new run and returns its full path
function New-RunFolder {
    $resultsRoot = Get-ResultsRoot
    $date = Get-Date -Format "yyyy-MM-dd"

    $candidate = Join-Path -Path $resultsRoot -ChildPath $date
    $n = 1
    while (Test-Path $candidate) {
        $candidate = Join-Path -Path $resultsRoot -ChildPath "$date ($n)"
        $n++
    }

    New-Item -ItemType Directory -Path $candidate -Force | Out-Null
    return $candidate
}

# Full path of an existing run folder: the one named $Name, or the most recent one
function Get-RunFolder {
    param([string] $Name)

    $resultsRoot = Get-ResultsRoot

    if ($Name) {
        $folder = Join-Path -Path $resultsRoot -ChildPath $Name
        if (-not (Test-Path $folder)) { throw "No existe la carpeta de resultados '$folder'." }
        return $folder
    }

    $latest = Get-ChildItem -Path $resultsRoot -Directory -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $latest) { throw "No hay corridas anteriores en '$resultsRoot'." }
    return $latest.FullName
}

Export-ModuleMember -Function New-RunFolder, Get-RunFolder
