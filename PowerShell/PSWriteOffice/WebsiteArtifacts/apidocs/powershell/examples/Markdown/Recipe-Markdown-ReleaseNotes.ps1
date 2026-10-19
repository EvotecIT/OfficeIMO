$path = '.\Release-Notes-4.2.0.md'
$changes = @(
    [pscustomobject]@{ Area = 'Reports'; Change = 'Added weekly PDF summary'; Audience = 'Operators' }
    [pscustomobject]@{ Area = 'Excel'; Change = 'Improved workbook validation'; Audience = 'Automation owners' }
    [pscustomobject]@{ Area = 'PowerPoint'; Change = 'Added speaker-note templates'; Audience = 'Presenters' }
)

MarkdownNew -Path $path {
    MarkdownFrontMatter -Data @{ title = 'Release 4.2.0'; date = '2026-08-18'; tags = @('release', 'automation') }
    MarkdownHeading -Level 1 -Text 'Release 4.2.0'
    MarkdownParagraph -Text 'This release adds report delivery options and tighter validation for generated files.'
    MarkdownCallout -Kind tip -Title 'Upgrade' -Body 'Test the release against one representative report before changing the production pin.'

    MarkdownHeading -Level 2 -Text 'What changed'
    MarkdownTable -InputObject $changes

    MarkdownHeading -Level 2 -Text 'Upgrade checklist'
    MarkdownTaskList -Items 'Update the module pin', 'Run the representative report', 'Inspect the generated files', 'Publish the approved version'

    MarkdownHeading -Level 2 -Text 'Install'
    MarkdownCode -Language powershell -Content "Install-Module ExampleModule -RequiredVersion 4.2.0 -Scope CurrentUser"

    MarkdownHeading -Level 2 -Text 'Known limits'
    MarkdownList -Items 'Existing templates are not modified automatically.', 'PDF signatures must be applied after content generation.'
}
