param([string] $OutputDirectory = (Join-Path $PSScriptRoot '../Documents'))
Import-Module PSWriteOffice -ErrorAction Stop
$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$path = Join-Path $OutputDirectory 'PowerPoint-Updated-Existing.pptx'
New-OfficePowerPoint -Path $path {
    PowerPointSlide -Title 'FY24 results' {
        PowerPointParagraph 'FY24 revenue is ready.'
        PowerPointNotes 'Explain the FY24 result.'
    }
    PowerPointSlide -Title 'FY24 priorities' {
        PowerPointList 'Retain customers', 'Improve margin'
    }
}
Open-OfficePowerPoint -Path $path {
    Update-OfficePowerPointText 'FY24' 'FY25' -IncludeNotes
}
Write-Host "Presentation saved to $path"
