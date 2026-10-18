Set-StrictMode -Version Latest

$helperRoot = $PSScriptRoot
. (Join-Path $helperRoot 'excel-performance.core.ps1')
. (Join-Path $helperRoot 'excel-performance.data.ps1')
. (Join-Path $helperRoot 'excel-performance.read.operations.ps1')
. (Join-Path $helperRoot 'excel-performance.operations.ps1')
. (Join-Path $helperRoot 'excel-performance.validation.ps1')
