param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Visio'),
    [switch] $Open
)

$ErrorActionPreference = 'Stop'

Import-Module PSWriteOffice -ErrorAction Stop

$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$path = Join-Path $OutputDirectory 'Example-Visio-ArchitectureMap.vsdx'
$svgPath = Join-Path $OutputDirectory 'Example-Visio-ArchitectureMap.svg'
$pngPath = Join-Path $OutputDirectory 'Example-Visio-ArchitectureMap.png'

OfficeVisio -Path $path -Title 'Service architecture map' -Author 'PSWriteOffice' -Width 12 -Height 7.5 -UseMastersByDefault -RequestRecalcOnOpen {
    VisioStencilImport -BuiltIn Architecture -Name Arch -Default
    VisioStencilImport -BuiltIn Cloud -Name Cloud
    VisioStencilImport -BuiltIn SecurityIdentity -Name Security
    VisioStencilImport -BuiltIn DataPlatform -Name Data

    VisioTextBox 'SaaS control plane' -X 6 -Y 6.85 -Width 4.2 -Height 0.42 -FillColor '#FFFFFF' -LineColor '#FFFFFF'
    VisioTextBox 'Boundaries, trust points, and platform services are editable Visio shapes.' -X 6 -Y 6.43 -Width 6.8 -Height 0.28 -FillColor '#FFFFFF' -LineColor '#FFFFFF'

    VisioRectangle -Key public -X 2.05 -Y 3.62 -Width 3.45 -Height 4.55 -FillColor '#F8FAFC' -LineColor '#CBD5E1'
    VisioRectangle -Key private -X 6.25 -Y 3.62 -Width 4.5 -Height 4.55 -FillColor '#EFF6FF' -LineColor '#93C5FD'
    VisioRectangle -Key datazone -X 10.1 -Y 3.62 -Width 2.4 -Height 4.55 -FillColor '#F0FDFA' -LineColor '#5EEAD4'
    VisioTextBox 'Edge' -X 2.05 -Y 5.65 -Width 1.8 -Height 0.3 -FillColor '#F8FAFC' -LineColor '#F8FAFC'
    VisioTextBox 'Application' -X 6.25 -Y 5.65 -Width 2.3 -Height 0.3 -FillColor '#EFF6FF' -LineColor '#EFF6FF'
    VisioTextBox 'Data' -X 10.1 -Y 5.65 -Width 1.8 -Height 0.3 -FillColor '#F0FDFA' -LineColor '#F0FDFA'

    VisioStencil -Catalog Arch -Stencil actor -Key user -Text 'Users' -X 0.8 -Y 3.8 -FillColor '#E0F2FE' -LineColor '#0284C7'
    VisioStencil -Catalog Cloud -Stencil gateway -Key gateway -Text 'API gateway' -X 2.4 -Y 4.9 -FillColor '#DBEAFE' -LineColor '#2563EB'
    VisioStencil -Catalog Security -Stencil policy -Key policy -Text 'Policy' -X 2.4 -Y 2.7 -FillColor '#FEF3C7' -LineColor '#D97706'
    VisioStencil -Catalog Arch -Stencil service -Key api -Text 'Public API' -X 4.6 -Y 4.9 -FillColor '#E0E7FF' -LineColor '#4F46E5'
    VisioStencil -Catalog Cloud -Stencil function -Key worker -Text 'Workers' -X 6.55 -Y 4.9 -FillColor '#DCFCE7' -LineColor '#16A34A'
    VisioStencil -Catalog Arch -Stencil queue -Key queue -Text 'Queue' -X 6.55 -Y 2.7 -FillColor '#CCFBF1' -LineColor '#0D9488'
    VisioStencil -Catalog Data -Stencil database -Key sql -Text 'Operational DB' -X 9.35 -Y 4.9 -FillColor '#F5F3FF' -LineColor '#7C3AED'
    VisioStencil -Catalog Arch -Stencil storage -Key archive -Text 'Archive' -X 9.35 -Y 2.7 -FillColor '#FAE8FF' -LineColor '#C026D3'
    VisioStencil -Catalog Cloud -Stencil monitoring -Key monitor -Text 'Telemetry' -X 11.05 -Y 5.7 -FillColor '#FFE4E6' -LineColor '#E11D48'

    VisioConnector -From user -To gateway -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -LineColor '#2563EB'
    VisioConnector -From gateway -To api -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -Label 'HTTPS' -LineColor '#2563EB'
    VisioConnector -From policy -To gateway -Kind Straight -FromSide Top -ToSide Bottom -EndArrow Triangle -LineColor '#D97706'
    VisioConnector -From api -To worker -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -Label 'commands' -LineColor '#16A34A'
    VisioConnector -From worker -To sql -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -Label 'write' -LineColor '#7C3AED'
    VisioConnector -From worker -To queue -Kind Straight -FromSide Bottom -ToSide Top -EndArrow Triangle -Label 'async' -LineColor '#0D9488'
    VisioConnector -From queue -To archive -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -LineColor '#C026D3'
    VisioConnector -From sql -To monitor -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -LineColor '#E11D48'
}

ConvertTo-OfficeVisioSvg -Path $path -OutputPath $svgPath
ConvertTo-OfficeVisioPng -Path $path -OutputPath $pngPath

if ($Open) {
    Invoke-Item $svgPath
}

[pscustomobject]@{
    Name = 'Service architecture map'
    Vsdx = $path
    Svg  = $svgPath
    Png  = $pngPath
}
