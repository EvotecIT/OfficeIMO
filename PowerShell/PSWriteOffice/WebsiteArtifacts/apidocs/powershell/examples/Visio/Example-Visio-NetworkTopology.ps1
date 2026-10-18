param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Visio'),
    [switch] $Open
)

$ErrorActionPreference = 'Stop'

Import-Module PSWriteOffice -ErrorAction Stop

$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$path = Join-Path $OutputDirectory 'Example-Visio-NetworkTopology.vsdx'
$svgPath = Join-Path $OutputDirectory 'Example-Visio-NetworkTopology.svg'
$pngPath = Join-Path $OutputDirectory 'Example-Visio-NetworkTopology.png'

OfficeVisio -Path $path -Title 'Branch office topology' -Author 'PSWriteOffice' -Width 11 -Height 7 -UseMastersByDefault -RequestRecalcOnOpen {
    VisioStencilImport -BuiltIn Network -Name Net -Default
    VisioStencilImport -BuiltIn Infrastructure -Name Infra

    VisioTextBox 'Branch office network' -X 5.5 -Y 6.35 -Width 4.5 -Height 0.42 -FillColor '#FFFFFF' -LineColor '#FFFFFF'
    VisioTextBox 'Zones, devices, and traffic paths from the OfficeIMO network stencil catalog.' -X 5.5 -Y 5.98 -Width 6.4 -Height 0.28 -FillColor '#FFFFFF' -LineColor '#FFFFFF'

    VisioRectangle -Key wanZone -X 1.65 -Y 3.2 -Width 2.45 -Height 4.1 -FillColor '#EFF6FF' -LineColor '#93C5FD'
    VisioRectangle -Key lanZone -X 5.25 -Y 3.2 -Width 4.1 -Height 4.1 -FillColor '#F0FDFA' -LineColor '#5EEAD4'
    VisioRectangle -Key serviceZone -X 9.2 -Y 3.2 -Width 2.2 -Height 4.1 -FillColor '#FDF2F8' -LineColor '#F9A8D4'
    VisioTextBox 'WAN' -X 1.65 -Y 4.95 -Width 1.5 -Height 0.28 -FillColor '#EFF6FF' -LineColor '#EFF6FF'
    VisioTextBox 'LAN' -X 5.25 -Y 4.95 -Width 1.5 -Height 0.28 -FillColor '#F0FDFA' -LineColor '#F0FDFA'
    VisioTextBox 'Services' -X 9.2 -Y 4.95 -Width 1.5 -Height 0.28 -FillColor '#FDF2F8' -LineColor '#FDF2F8'

    VisioStencil -Catalog Net -Stencil internet -Key internet -Text 'Internet' -X 0.85 -Y 3.45 -FillColor '#DBEAFE' -LineColor '#2563EB'
    VisioStencil -Catalog Net -Stencil firewall -Key firewall -Text 'Firewall' -X 2.55 -Y 3.45 -FillColor '#FEF3C7' -LineColor '#D97706'
    VisioStencil -Catalog Net -Stencil switch -Key core -Text 'Core switch' -X 4.45 -Y 3.45 -FillColor '#CCFBF1' -LineColor '#0D9488'
    VisioStencil -Catalog Net -Stencil wireless -Key wifi -Text 'Wi-Fi' -X 4.45 -Y 1.85 -FillColor '#E0F2FE' -LineColor '#0284C7'
    VisioStencil -Catalog Net -Stencil workstation -Key endpoints -Text 'Endpoints' -X 6.25 -Y 1.85 -FillColor '#DCFCE7' -LineColor '#16A34A'
    VisioStencil -Catalog Infra -Stencil server -Key app -Text 'App host' -X 7.1 -Y 4.35 -FillColor '#EDE9FE' -LineColor '#7C3AED'
    VisioStencil -Catalog Infra -Stencil storage-array -Key nas -Text 'NAS' -X 9.45 -Y 4.35 -FillColor '#FAE8FF' -LineColor '#C026D3'
    VisioStencil -Catalog Net -Stencil printer -Key printer -Text 'Printer' -X 6.25 -Y 3.45 -FillColor '#F8FAFC' -LineColor '#64748B'

    VisioConnector -From internet -To firewall -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -LineColor '#2563EB'
    VisioConnector -From firewall -To core -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -Label 'filtered' -LineColor '#D97706'
    VisioConnector -From core -To wifi -Kind Straight -FromSide Bottom -ToSide Top -EndArrow Triangle -LineColor '#0284C7'
    VisioConnector -From wifi -To endpoints -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -Label '802.1x' -LineColor '#16A34A'
    VisioConnector -From core -To printer -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -LineColor '#64748B'
    VisioConnector -From core -To app -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -Label 'VLAN 20' -LineColor '#7C3AED'
    VisioConnector -From app -To nas -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -Label 'backup' -LineColor '#C026D3'
}

ConvertTo-OfficeVisioSvg -Path $path -OutputPath $svgPath
ConvertTo-OfficeVisioPng -Path $path -OutputPath $pngPath

if ($Open) {
    Invoke-Item $svgPath
}

[pscustomobject]@{
    Name = 'Branch office topology'
    Vsdx = $path
    Svg  = $svgPath
    Png  = $pngPath
}
