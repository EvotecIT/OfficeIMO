param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Visio'),
    [switch] $Open
)

$ErrorActionPreference = 'Stop'

Import-Module PSWriteOffice -ErrorAction Stop

$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$path = Join-Path $OutputDirectory 'Example-Visio-StencilFlow.vsdx'
$svgPath = Join-Path $OutputDirectory 'Example-Visio-StencilFlow.svg'
$pngPath = Join-Path $OutputDirectory 'Example-Visio-StencilFlow.png'

New-OfficeVisio -Path $path -Title 'Customer onboarding flow' -Author 'PSWriteOffice' -Width 11in -Height 8.5in -UseMastersByDefault -RequestRecalcOnOpen {
    VisioStencilImport -BuiltIn Flowchart -Name Flow -Default

    VisioTextBox 'Customer onboarding' -X 5.5in -Y 7.55in -Width 5.2in -Height 0.42in -FillColor '#FFFFFF' -LineColor '#FFFFFF'
    VisioTextBox 'A compact, editable flowchart generated from PowerShell and OfficeIMO stencils.' -X 5.5in -Y 7.08in -Width 6.4in -Height 0.32in -FillColor '#FFFFFF' -LineColor '#FFFFFF'

    VisioRectangle -Key laneA -X 2.35in -Y 4.15in -Width 3.6in -Height 4.2in -FillColor '#EFF6FF' -LineColor '#BFDBFE'
    VisioRectangle -Key laneB -X 6.95in -Y 4.15in -Width 4.6in -Height 4.2in -FillColor '#F0FDFA' -LineColor '#99F6E4'
    VisioTextBox 'Intake' -X 2.35in -Y 6.05in -Width 2.8in -Height 0.32in -FillColor '#EFF6FF' -LineColor '#EFF6FF'
    VisioTextBox 'Automation' -X 6.95in -Y 6.05in -Width 3.2in -Height 0.32in -FillColor '#F0FDFA' -LineColor '#F0FDFA'

    VisioStencil -Catalog Flow -Stencil start -Key start -Text 'Request' -X 1.1 -Y 4.9 -FillColor '#DBEAFE' -LineColor '#2563EB'
    VisioStencil -Catalog Flow -Stencil process -Key validate -Text 'Validate profile' -X 3.0 -Y 4.9 -FillColor '#E0F2FE' -LineColor '#0284C7'
    VisioStencil -Catalog Flow -Stencil decision -Key decision -Text 'Complete?' -X 5.15 -Y 4.9 -Width 1.45 -Height 1.05 -FillColor '#FEF3C7' -LineColor '#D97706'
    VisioStencil -Catalog Flow -Stencil process -Key provision -Text 'Provision access' -X 7.35 -Y 5.65 -FillColor '#DCFCE7' -LineColor '#16A34A'
    VisioStencil -Catalog Flow -Stencil data -Key packet -Text 'Welcome packet' -X 9.55 -Y 5.65 -FillColor '#F5F3FF' -LineColor '#7C3AED'
    VisioStencil -Catalog Flow -Stencil process -Key rework -Text 'Collect missing data' -X 5.15 -Y 3.15 -FillColor '#FFE4E6' -LineColor '#E11D48'
    VisioStencil -Catalog Flow -Stencil end -Key done -Text 'Active customer' -X 9.55 -Y 3.15 -FillColor '#CCFBF1' -LineColor '#0F766E'

    VisioConnector -From start -To validate -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -LineColor '#2563EB'
    VisioConnector -From validate -To decision -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -LineColor '#2563EB'
    VisioConnector -From decision -To provision -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -Label 'yes' -LineColor '#16A34A'
    VisioConnector -From provision -To packet -Kind Straight -FromSide Right -ToSide Left -EndArrow Triangle -LineColor '#16A34A'
    VisioConnector -From packet -To done -Kind Straight -FromSide Bottom -ToSide Top -EndArrow Triangle -LineColor '#0F766E'
    VisioConnector -From decision -To rework -Kind Straight -FromSide Bottom -ToSide Top -EndArrow Triangle -Label 'no' -LineColor '#E11D48'
    VisioConnector -From rework -To validate -Kind Straight -FromSide Left -ToSide Bottom -EndArrow Triangle -LineColor '#E11D48'
}

Export-OfficeDocumentImage -Path $path -Page 1 -Format Svg -OutputPath $svgPath | Out-Null
Export-OfficeDocumentImage -Path $path -Page 1 -Format Png -OutputPath $pngPath | Out-Null

if ($Open) {
    Invoke-Item $svgPath
}

[pscustomobject]@{
    Name = 'Customer onboarding flow'
    Vsdx = $path
    Svg  = $svgPath
    Png  = $pngPath
}
