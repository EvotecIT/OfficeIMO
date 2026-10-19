$path = '.\Service-Restart-Runbook.md'
$checks = @(
    [pscustomobject]@{ Check = 'Health endpoint'; Expected = 'HTTP 200'; Owner = 'Operations' }
    [pscustomobject]@{ Check = 'Queue depth'; Expected = 'Below 100'; Owner = 'Application' }
    [pscustomobject]@{ Check = 'Error rate'; Expected = 'Below 1%'; Owner = 'Monitoring' }
)

MarkdownNew -Path $path {
    MarkdownFrontMatter -Data @{ title = 'Service restart runbook'; owner = 'Operations'; reviewed = '2026-08-18' }
    MarkdownTableOfContents -Title 'On this page' -PlaceAtTop -MinLevel 2 -MaxLevel 3
    MarkdownHeading -Level 1 -Text 'Service Restart Runbook'
    MarkdownCallout -Kind warning -Title 'Production change' -Body 'Confirm the approved change window before running any restart command.'

    MarkdownHeading -Level 2 -Text 'Before the restart'
    MarkdownTaskList -Items 'Confirm incident or change record', 'Notify the service owner', 'Capture current health metrics', 'Verify a rollback path'

    MarkdownHeading -Level 2 -Text 'Restart'
    MarkdownCode -Language powershell -Content "Restart-Service -Name 'ExampleService' -PassThru`nGet-Service -Name 'ExampleService'"

    MarkdownHeading -Level 2 -Text 'Validate'
    MarkdownTable -InputObject $checks
    MarkdownCallout -Kind note -Title 'Evidence' -Body 'Attach command output and health screenshots to the change record.'

    MarkdownHeading -Level 2 -Text 'Rollback'
    MarkdownDetails -Summary 'Show rollback steps' {
        MarkdownList -Items 'Stop the new service version', 'Restore the previous configuration', 'Start the previous version', 'Repeat validation checks'
    }
}
