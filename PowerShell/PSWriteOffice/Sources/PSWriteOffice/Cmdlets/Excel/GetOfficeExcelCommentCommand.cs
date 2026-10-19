using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets legacy worksheet comments (notes) from one or more worksheets.</summary>
/// <example>
///   <summary>Find comments containing review text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$comments = Get-OfficeExcelComment -Path .\Report.xlsx -Sheet Data -TextContains review
/// $comments |
///     Select-Object SheetName, Address, Author, Text</code>
///   <para>Returns matching comment metadata without modifying the workbook.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelComment", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelComments")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelCommentCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to inspect.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to inspect outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name to inspect. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) to inspect. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>A1 cell address to match.</summary>
    [Parameter]
    public string? Address { get; set; }

    /// <summary>A1 cell or range to match.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Comment author to match, ignoring case.</summary>
    [Parameter]
    public string? Author { get; set; }

    /// <summary>Text fragment to match, ignoring case.</summary>
    [Parameter]
    public string? TextContains { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: true);
        var path = string.Equals(ParameterSetName, ParameterSetPath, StringComparison.OrdinalIgnoreCase)
            ? Path
            : null;
        var filter = CreateFilter();

        foreach (var sheet in ExcelWorkbookCommandService.ResolveSheets(this, workbook.Document, ParameterSetName, Sheet, SheetIndex)) {
            foreach (var comment in sheet.FindComments(filter)) {
                WriteObject(ExcelCommentRecordService.CreateRecord(comment, sheet.Name, path));
            }
        }
    }

    private ExcelCommentFilter? CreateFilter() {
        if (!string.IsNullOrWhiteSpace(Address) && !string.IsNullOrWhiteSpace(Range)) {
            throw new PSArgumentException("Specify either -Address or -Range, not both.");
        }

        if (string.IsNullOrWhiteSpace(Address) && string.IsNullOrWhiteSpace(Range) && string.IsNullOrWhiteSpace(Author) && string.IsNullOrWhiteSpace(TextContains)) {
            return null;
        }

        return new ExcelCommentFilter {
            A1Range = !string.IsNullOrWhiteSpace(Address) ? Address : Range,
            Author = Author,
            TextContains = TextContains
        };
    }
}