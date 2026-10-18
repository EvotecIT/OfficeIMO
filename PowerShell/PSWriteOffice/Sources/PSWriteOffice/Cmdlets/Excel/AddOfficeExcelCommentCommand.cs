using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds or replaces a comment (note) on a worksheet cell.</summary>
/// <example>
///   <summary>Add a comment to A1.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelComment -Address 'A1' -Text 'Review this cell' }</code>
///   <para>Creates a cell comment at A1.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelComment", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelComment")]
public sealed class AddOfficeExcelCommentCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public int? SheetIndex { get; set; }

    /// <summary>1-based row index.</summary>
    [Parameter]
    public int? Row { get; set; }

    /// <summary>1-based column index.</summary>
    [Parameter]
    public int? Column { get; set; }

    /// <summary>A1-style cell address (e.g., A1, C5).</summary>
    [Parameter]
    public string? Address { get; set; }

    /// <summary>Comment text.</summary>
    [Parameter(Mandatory = true)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Author name (optional).</summary>
    [Parameter]
    public string? Author { get; set; }

    /// <summary>Author initials (optional).</summary>
    [Parameter]
    public string? Initials { get; set; }

    /// <summary>Emit the worksheet after adding the comment.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();
        var (row, column) = ExcelHostExtensions.ResolveCellAddress(Row, Column, Address);
        var author = string.IsNullOrWhiteSpace(Author) ? Environment.UserName : Author!;
        sheet.SetComment(row, column, Text, author, Initials);

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            if (Document == null)
            {
                throw new PSArgumentException("Provide an Excel document.");
            }

            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        var context = ExcelDslContext.Require(this);
        return context.RequireSheet();
    }
}
