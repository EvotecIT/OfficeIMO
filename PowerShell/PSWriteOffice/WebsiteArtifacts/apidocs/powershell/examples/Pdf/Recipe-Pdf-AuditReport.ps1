$path = '.\Access-Audit-Report.pdf'
$findings = @(
    [pscustomobject]@{ Id = 'A-01'; Severity = 'High'; Finding = 'Dormant privileged accounts'; Owner = 'Identity'; Due = '2026-08-28' }
    [pscustomobject]@{ Id = 'A-02'; Severity = 'Medium'; Finding = 'Missing quarterly owner review'; Owner = 'Governance'; Due = '2026-09-04' }
    [pscustomobject]@{ Id = 'A-03'; Severity = 'Low'; Finding = 'Inconsistent evidence naming'; Owner = 'Operations'; Due = '2026-09-11' }
)

PdfNew -Path $path {
    PdfTheme Report
    PdfMetadata -Title 'Quarterly access audit' -Author 'Internal Audit' -Subject 'Synthetic access review example'
    PdfPageSetup -PageSize A4 -Margin 40
    PdfHeader 'Quarterly Access Audit'
    PdfFooter 'Internal | Page {page}/{pages}'
    PdfPageBorder -Color '#334155' -Width 0.8 -Inset 20

    PdfBookmark 'summary'
    PdfHeading 'Quarterly Access Audit' -Level 1
    PdfPanel 'Overall result: remediation required. One high-severity finding must close before the next review.'

    PdfHeading 'Scope and method' -Level 2
    PdfList -Items 'Privileged directory roles', 'Dormant accounts', 'Quarterly owner attestations', 'Evidence retention'

    PdfHeading 'Findings' -Level 2
    PdfTable -InputObject $findings -Property Id,Severity,Finding,Owner,Due -HeaderFill '#334155' -HeaderTextColor '#FFFFFF' -RowStripeFill '#F8FAFC' -AutoFitColumns -KeepWithNext

    PdfPageBreak
    PdfBookmark 'actions'
    PdfHeading 'Remediation plan' -Level 1
    foreach ($finding in $findings) {
        PdfHeading "$($finding.Id): $($finding.Finding)" -Level 2
        PdfText -Run @{
            Text = 'Owner: ', $finding.Owner, '    Due: ', $finding.Due, '    Severity: ', $finding.Severity
            Bold = $true, $false, $true, $false, $true, $false
        }
        PdfFormField -Name "response-$($finding.Id)" -Type Text -Value 'Record the agreed action and evidence location.' -Width 480 -Height 42
    }

    PdfHeading 'Approval' -Level 2
    PdfFormField -Name 'audit-owner' -Type Text -Value 'Audit owner' -Width 230
    PdfFormField -Name 'review-status' -Type Choice -Options 'Draft', 'Ready for review', 'Approved' -Value 'Draft' -Width 230
}
