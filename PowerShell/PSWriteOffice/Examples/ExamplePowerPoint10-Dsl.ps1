Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'DslExample.pptx'
$data = @(
    [pscustomobject]@{ Item = 'Alpha'; Qty = 10 }
    [pscustomobject]@{ Item = 'Beta'; Qty = 20 }
)

New-OfficePowerPoint -Path $path {
    PptSlide {
        PptTitle -Title 'Status Update'
        PptTextBox -Text 'Generated with PSWriteOffice' -X 80 -Y 150 -Width 360 -Height 60
        PptBullets -Bullets 'Wins','Risks','Next Steps' -X 430 -Y 150 -Width 260 -Height 200
        PptNotes -Text 'Keep this under five minutes.'
    }

    PptSlide {
        PptTitle -Title 'Inventory'
        PptTable -Data $data -X 60 -Y 160 -Width 420 -Height 200
    }
}

Write-Host "Presentation saved to $path"
