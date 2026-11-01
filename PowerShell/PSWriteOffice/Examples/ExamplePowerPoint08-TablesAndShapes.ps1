param([string] $OutputDirectory = (Join-Path $PSScriptRoot 'Documents'))

Import-Module PSWriteOffice -ErrorAction Stop
$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$path = Join-Path $OutputDirectory 'ExamplePowerPoint8-TablesAndShapes.pptx'
$data = @(
    [pscustomobject]@{ Product = 'Alpha'; Qty = 12; Revenue = 1200 }
    [pscustomobject]@{ Product = 'Beta'; Qty = 7; Revenue = 940 }
    [pscustomobject]@{ Product = 'Gamma'; Qty = 20; Revenue = 1840 }
)

New-OfficePowerPoint -Path $path -Force {
    PowerPointSlide -Title 'Tables & Shapes' {
        PowerPointTable -InputObject $data -X 60 -Y 140 -Width 420 -Height 200
        PowerPointShape -Type Ellipse -X 520 -Y 140 -Width 140 -Height 140 -FillColor '#FFE699' -OutlineColor '#C65911' -OutlineWidth 1
        PowerPointTextBox 'Highlights' -X 530 -Y 300 -Width 120 -Height 40
    }
}
Write-Host "Presentation saved to $path"
