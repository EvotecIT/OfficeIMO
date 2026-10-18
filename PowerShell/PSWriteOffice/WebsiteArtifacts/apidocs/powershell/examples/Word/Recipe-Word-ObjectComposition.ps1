$path = '.\Word-Object-Composition.docx'
$findings = @(
    [pscustomobject]@{ Finding = 'Dormant administrator account'; Owner = 'Identity'; Status = 'Open' }
    [pscustomobject]@{ Finding = 'Missing evidence link'; Owner = 'Operations'; Status = 'Resolved' }
)

$document = New-OfficeWord -Path $path -NoSave
$heading = $document | Add-OfficeWordParagraph -Text 'Access review' -Style Heading1 -PassThru
$heading | Add-OfficeWordText -Text ' — weekly summary' -Color '#475569'

$summary = $document | Add-OfficeWordParagraph -PassThru
$summary | Add-OfficeWordText -Run @{
    Text  = 'Owner: ', 'Security', '    Status: ', 'Review required'
    Bold  = $true, $false, $true, $true
    Color = $null, $null, $null, 'Crimson'
}

Add-OfficeWordTable -Document $document -InputObject $findings -Style GridTable4Accent1 -Layout AutoFitToWindow
$document | Save-OfficeWord
$document | Close-OfficeWord
