param(
    [string] $OutputDirectory = (Join-Path $PWD 'PowerPoint-SectionsAndImport')
)

Import-Module PSWriteOffice -ErrorAction Stop
$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$sourcePath = Join-Path $OutputDirectory 'PowerPoint-Source.pptx'
$targetPath = Join-Path $OutputDirectory 'PowerPoint-SectionsAndImport.pptx'

New-OfficePowerPoint -Path $sourcePath -Force {
    PowerPointSlide -Title 'FY24 Imported' {
        PowerPointParagraph 'FY24 details from source deck'
        PowerPointNotes 'FY24 source notes'
    }
}

New-OfficePowerPoint -Path $targetPath -Force -PassThru {
    PowerPointSlide -Title 'FY24 Overview' {
        PowerPointParagraph 'FY24 summary for leadership'
    }
    PowerPointSlide -Title 'FY24 Results'
    Add-OfficePowerPointSection -Name 'Intro' -Slide 1
    Add-OfficePowerPointSection -Name 'Results' -Slide 2
    Rename-OfficePowerPointSection -Name 'Results' -NewName 'Deep Dive'
    Update-OfficePowerPointText -OldValue 'FY24' -NewValue 'FY25'
    Copy-OfficePowerPointSlide -Slide 1 -Position 2
    Import-OfficePowerPointSlide -LiteralPath $sourcePath -SourceSlide 1 -Position 2
}
