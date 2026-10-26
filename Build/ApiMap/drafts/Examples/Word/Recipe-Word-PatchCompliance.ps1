<#
.SYNOPSIS
Builds a Word patch compliance report and exports it to PDF.

.DESCRIPTION
Creates a themed Word report from PowerShell objects, then converts it to PDF.
Spacing, table padding and heading styles come from the theme.

.EXAMPLE
pwsh -NoProfile -File .\Examples\Word\Recipe-Word-PatchCompliance.ps1 -OutputDirectory .\out
#>
#Requires -Modules @{ ModuleName = 'PSWriteOffice'; ModuleVersion = '3.6.0' }
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $OutputDirectory
)

$ErrorActionPreference = 'Stop'
$null = New-Item -ItemType Directory -Path $OutputDirectory -Force
$docxPath = Join-Path $OutputDirectory 'PatchComplianceReport.docx'
$pdfPath = Join-Path $OutputDirectory 'PatchComplianceReport.pdf'

#region Excerpt:pswriteoffice-patch-report
# Sample inventory; replace it with your own query. Any objects with properties work.
$sites = @(
    [pscustomobject]@{ Site = 'Warsaw';    Servers = 84; Patched = 80; Pending = 4; Status = 'On target' }
    [pscustomobject]@{ Site = 'Frankfurt'; Servers = 66; Patched = 58; Pending = 8; Status = 'Watch' }
    [pscustomobject]@{ Site = 'Dublin';    Servers = 52; Patched = 50; Pending = 2; Status = 'On target' }
    [pscustomobject]@{ Site = 'Chicago';   Servers = 71; Patched = 69; Pending = 2; Status = 'On target' }
    [pscustomobject]@{ Site = 'Singapore'; Servers = 38; Patched = 33; Pending = 5; Status = 'Watch' }
)
$servers = ($sites | Measure-Object Servers -Sum).Sum
$patched = ($sites | Measure-Object Patched -Sum).Sum
$rows = $sites | Select-Object Site, Servers, Patched, Pending,
    @{ Name = 'Compliance'; Expression = { "$([math]::Round(100 * $_.Patched / $_.Servers, 1))%" } }, Status

New-OfficeWord -Path $docxPath -Theme Modern -Force {
    WordHeader { WordParagraph 'Contoso IT operations - week 41' }
    WordFooter { WordPageNumber -IncludeTotal }

    WordHeading 'Patch compliance report'
    WordParagraph {
        WordText "$([math]::Round(100 * $patched / $servers, 1))%" -Bold -Size 20 -Color Success
        WordText "  of $servers production servers are patched, $($servers - $patched) are pending."
    }

    WordHeading 'Compliance by site'
    WordTable $rows {
        WordTableCondition { $_.Status -eq 'Watch' } -BackgroundColor Warning
    }
    WordChart -Type Bar -InputObject $rows -Category Site -Value Patched, Pending -Title 'Patched and pending by site' -Legend -Height 7cm

    WordHeading 'Next steps'
    WordList -Style Numbered {
        WordListItem 'Frankfurt: patch the eight pending servers in the Thursday window.'
        WordListItem 'Singapore: confirm the maintenance slot with the regional owner.'
    }
}
Convert-OfficeDocument -Path $docxPath -OutputPath $pdfPath | Out-Null
#endregion

Get-Item -LiteralPath $docxPath, $pdfPath | Select-Object Name, Length
