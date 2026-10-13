# Wetcom header style shared by the raw Excel and the annex: bold white text on green.
function Set-HeaderStyle {
    param([Parameter(Mandatory = $true)] $Worksheet)

    if ($null -eq $Worksheet.Dimension) { return }

    $headerRange = $Worksheet.Cells[1, 1, 1, $Worksheet.Dimension.End.Column]
    $headerRange.Style.Font.Bold = $true
    $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $headerRange.Style.Fill.PatternType = 'Solid'
    $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(0, 176, 80))
}

Export-ModuleMember -Function Set-HeaderStyle
