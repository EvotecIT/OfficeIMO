#pragma warning disable CS1591
using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a threaded comment or reply to an Excel worksheet.</summary>
/// <example>
///   <summary>Add a threaded review note and audit it before sharing.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$comment = Add-OfficeExcelThreadedComment -Path .\Review.xlsx -Sheet Data -Address C5 -Text 'Please confirm this variance.' -Author 'Finance Reviewer' -PassThru
/// Add-OfficeExcelThreadedComment -Path .\Review.xlsx -Sheet Data -Address C5 -Text 'Confirmed with sales ops.' -Author 'Report Owner' -ParentId $comment.Id
/// Get-OfficeExcelCommentAudit -Path .\Review.xlsx -IncludeComments |
///     Select-Object -ExpandProperty ThreadedComments</code>
///   <para>Uses OfficeIMO threaded-comment metadata authoring, including workbook person metadata, and keeps legacy notes separate.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelThreadedComment", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelThreadedComment")]
[OutputType(typeof(PSObject))]
public sealed class AddOfficeExcelThreadedCommentCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using path or document input.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using path or document input.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public int? SheetIndex { get; set; }

    /// <summary>A1-style cell address, such as C5.</summary>
    [Parameter(Mandatory = true)]
    public string Address { get; set; } = string.Empty;

    /// <summary>Threaded comment text.</summary>
    [Parameter(Mandatory = true)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Author display name stored in workbook person metadata.</summary>
    [Parameter]
    public string? Author { get; set; }

    /// <summary>Optional parent threaded-comment id when adding a reply.</summary>
    [Parameter]
    public string? ParentId { get; set; }

    /// <summary>Optional stable threaded-comment id.</summary>
    [Parameter]
    public string? Id { get; set; }

    /// <summary>Optional timestamp for the threaded comment.</summary>
    [Parameter]
    public DateTime? Date { get; set; }

    /// <summary>Mark the threaded comment as done/resolved.</summary>
    [Parameter]
    public SwitchParameter Done { get; set; }

    /// <summary>Do not save when operating on a path-owned workbook.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    public SwitchParameter NoSave { get; set; }

    /// <summary>Emit threaded-comment metadata.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    protected override void ProcessRecord() {
        if (ParameterSetName == ParameterSetPath) {
            using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, null!, readOnly: false);
            if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Update Excel workbook")) {
                return;
            }

            var result = AddComment(ExcelSheetResolver.Resolve(workbook.Document, Sheet, SheetIndex));
            if (!NoSave.IsPresent) {
                workbook.SaveIfOwned();
            }

            WriteResult(result);
            return;
        }

        ExcelSheet sheet = ParameterSetName == ParameterSetDocument
            ? ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex)
            : ExcelDslContext.Require(this).RequireSheet();
        if (!ExcelShouldProcessService.ShouldProcessTarget(this, sheet.Name, "Add Excel threaded comment")) {
            return;
        }

        WriteResult(AddComment(sheet));
    }

    private ExcelThreadedCommentResult AddComment(ExcelSheet sheet) {
        return sheet.AddThreadedComment(new ExcelThreadedCommentOptions {
            Address = Address,
            Text = Text,
            Author = string.IsNullOrWhiteSpace(Author) ? Environment.UserName : Author!,
            ParentId = ParentId,
            Id = Id,
            Date = Date,
            Done = Done.IsPresent
        });
    }

    private void WriteResult(ExcelThreadedCommentResult result) {
        if (!PassThru.IsPresent) {
            return;
        }

        var output = new PSObject();
        output.Properties.Add(new PSNoteProperty("SheetName", result.SheetName));
        output.Properties.Add(new PSNoteProperty("CellReference", result.CellReference));
        output.Properties.Add(new PSNoteProperty("Id", result.Id));
        output.Properties.Add(new PSNoteProperty("PersonId", result.PersonId));
        output.Properties.Add(new PSNoteProperty("Author", result.Author));
        output.Properties.Add(new PSNoteProperty("IsReply", result.IsReply));
        output.Properties.Add(new PSNoteProperty("Done", result.Done));
        WriteObject(output);
    }
}
#pragma warning restore CS1591