using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a note or a threaded comment (or reply) to a cell.</summary>
/// <example>
/// <summary>Add a note to a cell</summary>
/// <code>Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
///     ExcelSheet 'Assumptions' {
///         ExcelComment B4 'Inflation from the October central bank projection.' -Author 'Finance'
///     }
/// }</code>
/// <para>Explains the source of a figure in a note on B4.</para>
/// </example>
/// <example>
/// <summary>Start a review thread</summary>
/// <code>Open-OfficeExcel -Path .\Finance\Forecast-2027.xlsx {
///     ExcelSheet 'Forecast' {
///         ExcelComment F12 'Can we confirm the Q2 hiring plan before this goes to the board?' -Threaded -Author 'Anna Kowalska'
///     }
/// }</code>
/// <para>Adds a threaded comment asking for confirmation of a forecast figure, attributed to the reviewer.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelComment", DefaultParameterSetName = "Note")]
[Alias("ExcelComment")]
public sealed class AddOfficeExcelCommentCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>A1 address of the target cell.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Address { get; set; } = null!;

    /// <summary>Comment text.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Text { get; set; } = null!;

    /// <summary>Author name (optional).</summary>
    [Parameter()]
    public string? Author { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Author initials, supplied together with Author.</summary>
    [Parameter(ParameterSetName = "Note")]
    public string? Initials { get; set; }

    /// <summary>Creates a threaded comment instead of a note.</summary>
    [Parameter(ParameterSetName = "Threaded", Mandatory = true)]
    public SwitchParameter Threaded { get; set; }

    /// <summary>Optional parent threaded-comment id when adding a reply.</summary>
    [Parameter(ParameterSetName = "Threaded")]
    public string? ParentId { get; set; }

    /// <summary>Optional stable threaded-comment id.</summary>
    [Parameter(ParameterSetName = "Threaded")]
    public string? Id { get; set; }

    /// <summary>Optional timestamp for the threaded comment.</summary>
    [Parameter(ParameterSetName = "Threaded")]
    public DateTime? Date { get; set; }

    /// <summary>Mark the threaded comment as done/resolved.</summary>
    [Parameter(ParameterSetName = "Threaded")]
    public SwitchParameter Done { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        string author = Author ?? Environment.UserName;
        if (Threaded) {
            var comment = sheet.AddThreadedComment(new ExcelThreadedCommentOptions { Address = Address, Text = Text, Author = author, ParentId = ParentId, Id = Id, Date = Date, Done = Done });
            if (PassThru) WriteObject(comment);
        } else {
            var (row, column) = ExcelHostExtensions.ResolveCellAddress(null, null, Address);
            sheet.SetComment(row, column, Text, author, Initials);
            if (PassThru) WriteObject(ExcelCommentRecordService.CreateRecord(sheet.FindComments(new ExcelCommentFilter { A1Range = Address }).Single(), sheet.Name, sheet.Document.FilePath));
        }
    }

}
