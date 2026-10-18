#pragma warning disable CS1591
using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Runs OfficeIMO workbook diagnostics and optional safe repairs.</summary>
/// <example>
///   <summary>Gate a generated workbook before sending it to a user.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doctor = Test-OfficeExcelWorkbook -Path .\Report.xlsx -SkipOpenXmlValidation
/// if (-not $doctor.Passed) {
///     $doctor.Issues |
///         Sort-Object Severity,Category,SheetName,Address |
///         Format-Table Severity,Category,SheetName,Address,Message,RepairAction
/// }</code>
///   <para>Returns OfficeIMO workbook diagnostics for defined names, formulas, tables, drawings, connections, and package validation.</para>
/// </example>
[Cmdlet(VerbsDiagnostic.Test, "OfficeExcelWorkbook", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[Alias("ExcelWorkbookDoctor", "ExcelDoctor")]
[OutputType(typeof(PSObject))]
public sealed class TestOfficeExcelWorkbookCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Repair duplicate, invalid, or broken defined names before reporting.</summary>
    [Parameter]
    public SwitchParameter RepairDefinedNames { get; set; }
    /// <summary>Skip Open XML validator checks.</summary>
    [Parameter]
    public SwitchParameter SkipOpenXmlValidation { get; set; }
    /// <summary>Return only a Boolean pass/fail value.</summary>
    [Parameter]
    public SwitchParameter Quiet { get; set; }

    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: !RepairDefinedNames.IsPresent);
        if (RepairDefinedNames.IsPresent &&
            !ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Repair Excel workbook diagnostics")) {
            return;
        }

        var report = workbook.Document.RunWorkbookDoctor(new ExcelWorkbookDoctorOptions {
            ValidateOpenXml = !SkipOpenXmlValidation.IsPresent,
            RepairDefinedNames = RepairDefinedNames.IsPresent
        });

        if (RepairDefinedNames.IsPresent) {
            workbook.SaveIfOwned();
        }

        if (Quiet.IsPresent) {
            WriteObject(!report.HasErrors);
            return;
        }

        var output = new PSObject();
        output.Properties.Add(new PSNoteProperty("Path", workbook.Document.FilePath));
        output.Properties.Add(new PSNoteProperty("Passed", !report.HasErrors));
        output.Properties.Add(new PSNoteProperty("HasWarnings", report.HasWarnings));
        output.Properties.Add(new PSNoteProperty("IssueCount", report.Issues.Count));
        output.Properties.Add(new PSNoteProperty("RepairedIssueCount", report.RepairedIssueCount));
        output.Properties.Add(new PSNoteProperty("Issues", report.Issues.Select(CreateIssue).ToArray()));
        WriteObject(output);
    }

    private static PSObject CreateIssue(ExcelWorkbookDiagnosticIssue issue) {
        var item = new PSObject();
        item.Properties.Add(new PSNoteProperty("Category", issue.Category));
        item.Properties.Add(new PSNoteProperty("Severity", issue.Severity.ToString()));
        item.Properties.Add(new PSNoteProperty("Message", issue.Message));
        item.Properties.Add(new PSNoteProperty("SheetName", issue.SheetName));
        item.Properties.Add(new PSNoteProperty("Address", issue.Address));
        item.Properties.Add(new PSNoteProperty("RepairAction", issue.RepairAction));
        return item;
    }
}
#pragma warning restore CS1591