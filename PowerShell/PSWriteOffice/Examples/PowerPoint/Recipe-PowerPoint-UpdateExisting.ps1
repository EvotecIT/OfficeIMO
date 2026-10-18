$path = '.\PowerPoint-Updated-Existing.pptx'
PptNew -Path $path {
    PptSlide {
        PptTitle -Title 'FY24 Results'
        PptTextBox -Text 'FY24 revenue is ready.' -X 80 -Y 150 -Width 500 -Height 60
        PptNotes -Text 'Explain the FY24 result.'
    }
    PptSlide {
        PptTitle -Title 'FY24 Priorities'
        PptBullets -Bullets 'Retain customers', 'Improve margin' -X 80 -Y 150 -Width 500 -Height 140
    }
}

$presentation = Get-OfficePowerPoint -Path $path
Update-OfficePowerPointText -Presentation $presentation -OldValue 'FY24' -NewValue 'FY25' -IncludeNotes
Close-OfficePowerPoint -Presentation $presentation -Save
