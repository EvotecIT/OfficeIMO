#pragma warning disable CS1591
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Audits legacy notes and threaded comments preserved in an Excel workbook.</summary>
/// <example>
///   <summary>Review comments before distributing a workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$audit = Get-OfficeExcelCommentAudit -Path .\ReviewWorkbook.xlsx -IncludeComments
/// $audit.Comments | Sort-Object SheetName,CellReference | Format-Table SheetName,CellReference,Author,Text
/// $audit.Issues | Format-Table Severity,Category,SheetName,Address,Message</code>
///   <para>Returns workbook-level note/threaded-comment counts, optional comment records, and metadata issues such as missing authors or orphaned threaded replies.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelCommentAudit", DefaultParameterSetName = ParameterSetPath)]
[Alias("ExcelCommentAudit", "ExcelCommentsAudit")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelCommentAuditCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Include legacy and threaded comment records in the output.</summary>
    [Parameter]
    public SwitchParameter IncludeComments { get; set; }

    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: true);
        var report = workbook.Document.InspectComments();
        var output = new PSObject();
        output.Properties.Add(new PSNoteProperty("Path", workbook.Document.FilePath));
        output.Properties.Add(new PSNoteProperty("HasComments", report.HasComments));
        output.Properties.Add(new PSNoteProperty("CommentCount", report.CommentCount));
        output.Properties.Add(new PSNoteProperty("ThreadedCommentCount", report.ThreadedCommentCount));
        output.Properties.Add(new PSNoteProperty("IssueCount", report.Issues.Count));
        output.Properties.Add(new PSNoteProperty("Issues", report.Issues.Select(CreateIssue).ToArray()));
        if (IncludeComments.IsPresent) {
            output.Properties.Add(new PSNoteProperty("Comments", report.Comments.Select(CreateComment).ToArray()));
            output.Properties.Add(new PSNoteProperty("ThreadedComments", report.ThreadedComments.Select(CreateThreadedComment).ToArray()));
        }

        WriteObject(output);
    }

    private static PSObject CreateIssue(ExcelWorkbookDiagnosticIssue issue) {
        var item = new PSObject();
        item.Properties.Add(new PSNoteProperty("Category", issue.Category));
        item.Properties.Add(new PSNoteProperty("Severity", issue.Severity.ToString()));
        item.Properties.Add(new PSNoteProperty("Message", issue.Message));
        item.Properties.Add(new PSNoteProperty("SheetName", issue.SheetName));
        item.Properties.Add(new PSNoteProperty("Address", issue.Address));
        return item;
    }

    private static PSObject CreateComment(ExcelCommentRecord comment) {
        var item = new PSObject();
        item.Properties.Add(new PSNoteProperty("SheetName", comment.SheetName));
        item.Properties.Add(new PSNoteProperty("CellReference", comment.CellReference));
        item.Properties.Add(new PSNoteProperty("Author", comment.Author));
        item.Properties.Add(new PSNoteProperty("Text", comment.Text));
        return item;
    }

    private static PSObject CreateThreadedComment(ExcelThreadedCommentSnapshot comment) {
        var item = new PSObject();
        item.Properties.Add(new PSNoteProperty("SheetName", comment.SheetName));
        item.Properties.Add(new PSNoteProperty("CellReference", comment.CellReference));
        item.Properties.Add(new PSNoteProperty("Id", comment.Id));
        item.Properties.Add(new PSNoteProperty("ParentId", comment.ParentId));
        item.Properties.Add(new PSNoteProperty("Author", comment.Author));
        item.Properties.Add(new PSNoteProperty("Text", comment.Text));
        item.Properties.Add(new PSNoteProperty("Date", comment.Date));
        item.Properties.Add(new PSNoteProperty("Done", comment.Done));
        return item;
    }
}
#pragma warning restore CS1591