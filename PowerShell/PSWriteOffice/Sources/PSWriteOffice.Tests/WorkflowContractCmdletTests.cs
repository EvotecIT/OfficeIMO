using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Project;
using OfficeIMO.Invoicing;
using OfficeIMO.Pdf;
using PSWriteOffice.Cmdlets.Project;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class WorkflowContractCmdletTests {
    [Fact]
    public void CsvMappedProjectTablesRetainIdentityAndExposeProjectionLoss() {
        using var runspace = OpenRunspace();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            $rows = 'ID,Task,Duration', '7,Design,480', '8,Build,960' | ConvertFrom-Csv
            $table = $rows | ConvertTo-OfficeProjectTable -Kind Tasks -ColumnMap @{ Uid='ID'; Name='Task'; DurationMinutes='Duration' }
            $project = $table | Import-OfficeProjectTable -Name Delivery -StartDate 2026-10-20
            try {
                $tasks = @($project | Get-OfficeProjectTask)
                if ($tasks[0].Uid -ne 7 -or $tasks[1].Uid -ne 8) { throw 'Identities changed' }
                $rejected = $false
                try { $project | Export-OfficeProjectTable -ErrorAction Stop } catch { $rejected = $true }
                if (-not $rejected) { throw 'Projection loss was accepted implicitly' }
                $project | Export-OfficeProjectTable -AllowLossyProjection
            } finally { $project | Close-OfficeProject }
            """);
        var output = shell.Invoke();
        AssertNoErrors(shell);
        var result = Assert.IsType<ProjectDataExportResult>(Assert.Single(output).BaseObject);
        Assert.NotEmpty(result.Notices);
        var tasks = Assert.Single(result.Tables, table => table.Kind == ProjectDataKind.Tasks);
        Assert.Equal(new[] { "7", "8" }, tasks.Table.Rows.Select(row => row[0]));
        Assert.Equal(new[] { "480", "960" }, tasks.Table.Rows.Select(row => row[7]));
    }

    [Theory]
    [InlineData(".pdf")]
    [InlineData(".html")]
    [InlineData(".docx")]
    [InlineData(".xlsx")]
    [InlineData(".pptx")]
    public void CurrentProjectViewExportsReportsAfterApplyingSchedule(string extension) {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-project-report-" + Guid.NewGuid().ToString("N") + extension);
        using var runspace = OpenRunspace();
        runspace.SessionStateProxy.SetVariable("outputPath", path);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            $project = New-OfficeProject -Name Delivery -StartDate '2026-10-20 08:00'
            try {
                $project | Add-OfficeProjectTask -Name Build -DurationDays 2
                $project | Invoke-OfficeProjectSchedule -Apply | Out-Null
                $view = $project | Get-OfficeProjectView -Kind Gantt
                $view | Export-OfficeProjectReport -Path $outputPath -WhatIf
                if (Test-Path -LiteralPath $outputPath) { throw 'WhatIf wrote a report' }
                $view | Export-OfficeProjectReport -Path $outputPath
                $view
            } finally { $project | Close-OfficeProject }
            """);
        try {
            var output = shell.Invoke();
            AssertNoErrors(shell);
            var view = Assert.IsType<ProjectView>(Assert.Single(output).BaseObject);
            Assert.Equal("Build", Assert.Single(view.Rows).Name);
            Assert.NotEmpty(view.Render());
            Assert.True(new FileInfo(path).Length > 100);
            if (extension == ".pdf") Assert.Contains("Build", PdfDocument.Load(path).Read().Text);
            else if (extension == ".html") Assert.Contains("Build", File.ReadAllText(path));
            else if (extension == ".xlsx") {
                using var workbook = OfficeIMO.Excel.ExcelDocument.Load(path);
                Assert.True(workbook.Sheets[0].TryGetCellText(2, 2, out string name));
                Assert.Equal("Build", name);
            } else if (extension == ".docx") {
                using var document = OfficeIMO.Word.WordDocument.Load(path);
                Assert.NotEmpty(document.Tables);
            } else {
                using var presentation = OfficeIMO.PowerPoint.PowerPointDocument.Load(path);
                Assert.NotEmpty(presentation.Slides);
            }
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void EditedInvoiceRecalculatesExplicitlyAndRendersItsCapturedXml() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-invoice-workflow-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("outputDirectory", directory);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                $seller = New-OfficeInvoiceParty -Name Seller -CountryCode DE -VatId DE123456789
                $buyer = New-OfficeInvoiceParty -Name Buyer -CountryCode DE
                $invoice = New-OfficeInvoice -Number INV-001 -IssueDate 2026-10-20 -Seller $seller -Buyer $buyer
                $invoice | Add-OfficeInvoiceLine -Name Consulting -Quantity 2 -UnitPrice 100 -TaxRate 19
                $xml = Join-Path $outputDirectory 'source.xml'
                $pdf = Join-Path $outputDirectory 'invoice.pdf'
                $invoice | Save-OfficeInvoice -Path $xml
                $invoice = Get-OfficeInvoice -Path $xml
                $invoice | Get-OfficeInvoiceLine -Id 1 | Set-OfficeInvoiceLine -Quantity 3
                $invoice | Set-OfficeInvoice -TaxCurrency PLN -TaxAmountInAccountingCurrency 200
                $invoice | Set-OfficeInvoice -RecalculateTotals
                if ($null -ne $invoice.TaxAmountInAccountingCurrency) { throw 'Stale accounting VAT survived recalculation' }
                $invoice | Set-OfficeInvoice -TaxCurrency $null
                $invoice | Export-OfficeInvoicePdf -Path $pdf -WhatIf
                if (Test-Path -LiteralPath $pdf) { throw 'WhatIf wrote a hybrid PDF' }
                $invoice | Export-OfficeInvoicePdf -Path $pdf
                Get-OfficePdfAttachment -Path $pdf
                """);
            var output = shell.Invoke();
            AssertNoErrors(shell);
            var attachment = Assert.IsType<PdfExtractedAttachment>(Assert.Single(output).BaseObject);
            var embedded = InvoiceParser.Read(attachment.Bytes);
            Assert.Equal(3m, Assert.Single(embedded.Invoice.Lines).Quantity);
            Assert.Contains("INV-001", PdfDocument.Load(Path.Combine(directory, "invoice.pdf")).Read().Text);
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void LiteralProviderInputsRejectNonFileSystemAndAcceptCustomDrive() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-provider-workflow-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("rootPath", directory);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                New-PSDrive -Name AuditFiles -PSProvider FileSystem -Root $rootPath | Out-Null
                try {
                    $book = New-OfficeEpub -Title Literal
                    $book | Add-OfficeEpubChapter -Title Opening -Text Safe
                    $book | Save-OfficeEpub -Path 'AuditFiles:/book[1].epub'
                    Get-Item -LiteralPath 'AuditFiles:/book[1].epub' | Get-OfficeEpub | Get-OfficeEpubChapter
                    $rejected = $false
                    try { Get-OfficeEpub -Path 'Env:/PATH' -ErrorAction Stop } catch { $rejected = $_.Exception.Message -like '*filesystem*' }
                    if (-not $rejected) { throw 'Non-filesystem provider was accepted' }
                    foreach ($command in 'Get-OfficeCsv', 'Import-OfficeCsv', 'Get-OfficeDocumentBatch') {
                        $rejected = $false
                        try { & $command -Path 'Env:/*' -ErrorAction Stop } catch { $rejected = $_.Exception.Message -like '*filesystem*' }
                        if (-not $rejected) { throw "$command accepted a non-filesystem discovery provider" }
                    }
                } finally { Remove-PSDrive AuditFiles }
                """);
            var output = shell.Invoke();
            AssertNoErrors(shell);
            Assert.Single(output);
            Assert.True(File.Exists(Path.Combine(directory, "book[1].epub")));
        } finally { Directory.Delete(directory, true); }
    }

    private static Runspace OpenRunspace() {
        var state = InitialSessionState.CreateDefault();
        state.ImportPSModule(new[] { typeof(NewOfficeProjectCommand).Assembly.Location });
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
    private static void AssertNoErrors(PowerShell shell) => Assert.Empty(shell.Streams.Error);
}
