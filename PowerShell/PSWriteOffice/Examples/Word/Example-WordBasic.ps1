Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$data = @(
    [PSCustomObject]@{ Region = 'North America'; Revenue = 125000; YoY = '12%' }
    [PSCustomObject]@{ Region = 'EMEA'; Revenue = 98000; YoY = '22%' }
    [PSCustomObject]@{ Region = 'APAC'; Revenue = 143000; YoY = '18%' }
)

$docPath = Join-Path $documents 'Word-BasicDocument.docx'

New-OfficeWord -Path $docPath {
    Add-OfficeWordSection {
        Add-OfficeWordHeader -Type Default {
            Add-OfficeWordParagraph -Text 'Quarterly Revenue Snapshot' -Style Heading2
            Add-OfficeWordPageNumber -IncludeTotalPages
        }

        Add-OfficeWordParagraph -Text 'Executive Summary' -Style Heading1
        Add-OfficeWordParagraph -Text 'Revenue accelerated in all regions with double-digit YoY momentum.'

        Add-OfficeWordList -Style 'Numbered' {
            Add-OfficeWordListItem -Text 'North America +12% YoY'
            Add-OfficeWordListItem -Text 'EMEA +22% YoY'
            Add-OfficeWordListItem -Text 'APAC +18% YoY'
        }

        Add-OfficeWordTable -InputObject $data -Style 'GridTable1LightAccent2' {
            Add-OfficeWordTableCondition -FilterScript { $_.Revenue -gt 100000 } -BackgroundColor '#e6fffb'
        }
    }
}

Write-Host "Document saved to $docPath"
