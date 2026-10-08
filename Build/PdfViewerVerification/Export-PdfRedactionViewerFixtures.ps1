param(
    [string] $BinaryRoot = (Join-Path $PSScriptRoot '../../OfficeIMO.Pdf.Benchmarks/bin/Release/net8.0'),
    [Parameter(Mandatory)] [string] $OutputRoot,
    [string] $InputPath,
    [string] $Pattern = 'private account [0-9]{3}'
)
$ErrorActionPreference = 'Stop'
$BinaryRoot = (Resolve-Path -LiteralPath $BinaryRoot).Path
foreach ($name in 'OfficeIMO.Core', 'OfficeIMO.Pdf', 'OfficeIMO.Pdf.Benchmarks') {
    [void] [Reflection.Assembly]::LoadFrom((Join-Path $BinaryRoot "$name.dll"))
}
if ($InputPath) {
    $workload = [OfficeIMO.Pdf.Benchmarks.PdfRedactionRuntimeWorkload]::new((Resolve-Path -LiteralPath $InputPath).Path, $Pattern)
    $workload.Execute()
    $workload.Export($OutputRoot)
} else {
    foreach ($rotation in 0, 90, 180, 270) {
        $workload = [OfficeIMO.Pdf.Benchmarks.PdfRedactionRuntimeWorkload]::new(2, $rotation)
        $workload.Execute()
        $workload.Export((Join-Path $OutputRoot "rotation-$rotation"))
    }
}
