$path = '.\Service-Restart-Runbook.md'
$checks = @(
    [pscustomobject]@{ Check = 'Health endpoint'; Expected = 'HTTP 200'; Owner = 'Operations' }
    [pscustomobject]@{ Check = 'Queue depth'; Expected = 'Below 100'; Owner = 'Application' }
    [pscustomobject]@{ Check = 'Error rate'; Expected = 'Below 1%'; Owner = 'Monitoring' }
)

New-OfficeMarkdown -Path $path -Force {
    MarkdownFrontMatter @{ title = 'Service restart runbook'; owner = 'Operations'; reviewed = '2026-08-18' }
    MarkdownTableOfContents -Title 'On this page' -Placement Top -MinLevel 2 -MaxLevel 3
    MarkdownHeading 'Service restart runbook'
    MarkdownCallout warning 'Production change' 'Confirm the approved change window before running any restart command.'

    MarkdownHeading 'Before the restart' -Level 2
    MarkdownTaskList 'Confirm incident or change record', 'Notify the service owner', 'Capture current health metrics', 'Verify a rollback path'

    MarkdownHeading 'Restart' -Level 2
    MarkdownCode powershell @'
Restart-Service -Name 'ExampleService' -PassThru
Get-Service -Name 'ExampleService'
'@

    MarkdownHeading 'Validate' -Level 2
    MarkdownTable $checks
    MarkdownCallout note 'Evidence' 'Attach command output and health screenshots to the change record.'

    MarkdownHeading 'Rollback' -Level 2
    MarkdownDetails 'Show rollback steps' {
        MarkdownList -Ordered 'Stop the new service version', 'Restore the previous configuration', 'Start the previous version', 'Repeat validation checks'
    }
}
