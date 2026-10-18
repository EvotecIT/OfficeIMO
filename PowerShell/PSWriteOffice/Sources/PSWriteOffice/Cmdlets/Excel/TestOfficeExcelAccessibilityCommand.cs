#pragma warning disable CS1591
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Checks workbook accessibility and compliance signals.</summary>
/// <example>
///   <summary>Review workbook accessibility signals before exporting a report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$accessibility = Test-OfficeExcelAccessibility -Path .\Dashboard.xlsx
/// $accessibility.Findings |
///     Sort-Object Severity,Category,SheetName,Address |
///     Format-Table Severity,Category,SheetName,Address,Message</code>
///   <para>Reports OfficeIMO accessibility findings such as missing image alt text, hidden sheets, merged ranges, and tables without header rows.</para>
/// </example>
[Cmdlet(VerbsDiagnostic.Test, "OfficeExcelAccessibility", DefaultParameterSetName = ParameterSetPath)]
[Alias("ExcelAccessibility")]
[OutputType(typeof(PSObject))]
public sealed class TestOfficeExcelAccessibilityCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Return only a Boolean pass/fail value.</summary>
    [Parameter]
    public SwitchParameter Quiet { get; set; }

    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: true);
        var report = workbook.Document.AnalyzeAccessibility();
        if (Quiet.IsPresent) {
            WriteObject(!report.HasWarnings);
            return;
        }

        var output = new PSObject();
        output.Properties.Add(new PSNoteProperty("Path", workbook.Document.FilePath));
        output.Properties.Add(new PSNoteProperty("Passed", !report.HasWarnings));
        output.Properties.Add(new PSNoteProperty("FindingCount", report.Findings.Count));
        output.Properties.Add(new PSNoteProperty("Findings", report.Findings.Select(CreateFinding).ToArray()));
        WriteObject(output);
    }

    private static PSObject CreateFinding(ExcelAccessibilityFinding finding) {
        var item = new PSObject();
        item.Properties.Add(new PSNoteProperty("Category", finding.Category));
        item.Properties.Add(new PSNoteProperty("Severity", finding.Severity.ToString()));
        item.Properties.Add(new PSNoteProperty("Message", finding.Message));
        item.Properties.Add(new PSNoteProperty("SheetName", finding.SheetName));
        item.Properties.Add(new PSNoteProperty("Address", finding.Address));
        return item;
    }
}
#pragma warning restore CS1591