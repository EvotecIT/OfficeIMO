param(
    [string] $StencilPackagePath,
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Visio'),
    [switch] $Open
)

$ErrorActionPreference = 'Stop'

Import-Module PSWriteOffice -ErrorAction Stop

$repoRoot = if ($env:EVOTEC_GITHUB_ROOT) {
    $env:EVOTEC_GITHUB_ROOT
} elseif ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
    'C:\Support\GitHub'
} else {
    Join-Path $HOME 'Support/GitHub'
}

if (-not $StencilPackagePath) {
    $StencilPackagePath = Join-Path $repoRoot 'OfficeIMO\Assets\VisioTemplates\DrawingWithShapes.vsdx'
}

if (-not (Test-Path -LiteralPath $StencilPackagePath)) {
    throw "Stencil package was not found. Provide -StencilPackagePath with a .vssx, .vstx, or .vsdx file."
}

$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$path = Join-Path $OutputDirectory 'Example-Visio-PackageStencil.vsdx'
$svgPath = Join-Path $OutputDirectory 'Example-Visio-PackageStencil.svg'
$pngPath = Join-Path $OutputDirectory 'Example-Visio-PackageStencil.png'

$catalog = Get-OfficeVisioStencilCatalog -Path $StencilPackagePath -CatalogName 'Loaded package' -IncludeUnsupportedMasters
$sampleStencils = @(Find-OfficeVisioStencil -Catalog $catalog -First 3)
if ($sampleStencils.Count -eq 0) {
    throw "No stencil masters were discovered in $StencilPackagePath."
}

OfficeVisio -Path $path -Title 'Package-backed stencils' -Author 'PSWriteOffice' -Width 10 -Height 6.5 -UseMastersByDefault -RequestRecalcOnOpen {
    VisioStencilImport -Catalog $catalog -Name Package -Default
    VisioTextBox 'Package-backed stencil import' -X 5 -Y 5.8 -Width 4.6 -Height 0.38 -FillColor '#FFFFFF' -LineColor '#FFFFFF'
    VisioTextBox "Loaded from $([System.IO.Path]::GetFileName($StencilPackagePath))" -X 5 -Y 5.42 -Width 5.2 -Height 0.26 -FillColor '#FFFFFF' -LineColor '#FFFFFF'

    $first = $sampleStencils[0].Id
    $second = if ($sampleStencils.Count -gt 1) { $sampleStencils[1].Id } else { $sampleStencils[0].Id }
    $third = if ($sampleStencils.Count -gt 2) { $sampleStencils[2].Id } else { $sampleStencils[0].Id }

    VisioStencil -Stencil $first -Key importedA -Text 'Package master A' -X 2 -Y 3.4 -FillColor '#E0F2FE' -LineColor '#0284C7'
    VisioStencil -Stencil $second -Key importedB -Text 'Package master B' -X 5 -Y 3.4 -FillColor '#FEF3C7' -LineColor '#D97706'
    VisioStencil -Stencil $third -Key importedC -Text 'Package master C' -X 8 -Y 3.4 -FillColor '#DCFCE7' -LineColor '#16A34A'
    VisioConnector -From importedA -To importedB -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -Label 'loaded' -LineColor '#0284C7'
    VisioConnector -From importedB -To importedC -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -Label 'reused' -LineColor '#16A34A'
}

ConvertTo-OfficeVisioSvg -Path $path -OutputPath $svgPath
ConvertTo-OfficeVisioPng -Path $path -OutputPath $pngPath

if ($Open) {
    Invoke-Item $svgPath
}

[pscustomobject]@{
    Name = 'Package-backed stencils'
    Package = $StencilPackagePath
    Vsdx = $path
    Svg  = $svgPath
    Png  = $pngPath
}
