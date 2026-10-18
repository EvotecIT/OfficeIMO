Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'ExamplePowerPoint8-TablesAndShapes.pptx'
$presentation = New-OfficePowerPoint -Path $path -NoSave

$slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Tables & Shapes'

$data = @(
    [PSCustomObject]@{ Product = 'Alpha'; Qty = 12; Revenue = 1200 }
    [PSCustomObject]@{ Product = 'Beta'; Qty = 7; Revenue = 940 }
    [PSCustomObject]@{ Product = 'Gamma'; Qty = 20; Revenue = 1840 }
)

Add-OfficePowerPointTable -Slide $slide -Data $data -X 60 -Y 140 -Width 420 -Height 200
Add-OfficePowerPointShape -Slide $slide -ShapeType Ellipse -X 520 -Y 140 -Width 140 -Height 140 -FillColor '#FFE699' -OutlineColor '#C65911' -OutlineWidth 1
Add-OfficePowerPointTextBox -Slide $slide -Text 'Highlights' -X 530 -Y 300 -Width 120 -Height 40

Save-OfficePowerPoint -Presentation $presentation
Write-Host "Presentation saved to $path"
