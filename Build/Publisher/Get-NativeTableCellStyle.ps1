function Get-NativeTableCellStyle {
    <#
    .SYNOPSIS
    Observes raw native cell fills and borders without changing them.
    .PARAMETER Cell
    An observed Publisher table cell.
    .PARAMETER Objects
    Retains observed COM wrappers for the isolated worker's lifetime.
    .OUTPUTS
    Independent fill and edge observations with explicit failures; border weights use the native property value.
    #>
    param($Cell, [Collections.Generic.List[object]] $Objects)

    try {
        $format = $Cell.Fill
        $color = $format.ForeColor
        $Objects.Add($format)
        $Objects.Add($color)
        $fill = [ordered]@{
            status = 'Observed'
            visible = [int] $format.Visible
            type = [int] $format.Type
            colorType = [int] $color.Type
            rgb = [int] $color.RGB
            transparency = [double] $format.Transparency
        }
    } catch {
        $fill = [ordered]@{ status = 'Unassessed'; errorHResult = ('0x{0:X8}' -f $_.Exception.HResult) }
    }
    $borders = [ordered]@{}
    foreach ($side in @('Left', 'Top', 'Right', 'Bottom', 'Diagonal')) {
        try {
            $format = $Cell.('Border' + $side)
            $color = $format.Color
            $Objects.Add($format)
            $Objects.Add($color)
            $borders[$side.ToLowerInvariant()] = [ordered]@{
                status = 'Observed'
                widthPoints = [double] $format.Weight
                colorType = [int] $color.Type
                rgb = [int] $color.RGB
            }
        } catch {
            $borders[$side.ToLowerInvariant()] = [ordered]@{ status = 'Unassessed'; errorHResult = ('0x{0:X8}' -f $_.Exception.HResult) }
        }
    }
    [pscustomobject]@{ fill = $fill; borders = $borders }
}
