using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Updates legacy worksheet comments (notes) that match a filter.</summary>
/// <example>
///   <summary>Update a comment on one cell.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$updated = Update-OfficeExcelComment -Path .\Report.xlsx -Sheet Data -Address B2 -Text 'Reviewed' -Author Carol -Initials CC -PassThru
/// Get-OfficeExcelComment -Path .\Report.xlsx -Sheet Data -Address B2 |
///     Select-Object Address, Author, Text</code>
///   <para>Replaces matching comment text and optionally changes the author.</para>
/// </example>
[Cmdlet(VerbsData.Update, "OfficeExcelComment", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Medium, DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelCommentUpdate")]
[OutputType(typeof(int))]
public sealed class UpdateOfficeExcelCommentCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name to update. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) to update. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>A1 cell address to match.</summary>
    [Parameter]
    public string? Address { get; set; }

    /// <summary>A1 cell or range to match.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Existing comment author to match, ignoring case.</summary>
    [Parameter]
    public string? MatchAuthor { get; set; }

    /// <summary>Existing text fragment to match, ignoring case.</summary>
    [Parameter]
    public string? TextContains { get; set; }

    /// <summary>Allow updating all comments on the selected worksheet(s) when no filter is supplied.</summary>
    [Parameter]
    public SwitchParameter All { get; set; }

    /// <summary>Replacement plain text.</summary>
    [Parameter]
    public string? Text { get; set; }

    /// <summary>Replacement rich text runs.</summary>
    [Parameter]
    [Alias("Runs")]
    public object[]? Run { get; set; }

    /// <summary>Replacement author name.</summary>
    [Parameter]
    public string? Author { get; set; }

    /// <summary>Replacement author initials.</summary>
    [Parameter]
    public string? Initials { get; set; }

    /// <summary>Returns the number of comments updated.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var filter = CreateRequiredFilter();
        if (string.IsNullOrWhiteSpace(Text) == (Run == null || Run.Length == 0)) {
            throw new PSArgumentException("Specify exactly one of -Text or -Run.");
        }

        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        var updated = 0;
        foreach (var sheet in ExcelWorkbookCommandService.ResolveSheets(this, workbook.Document, ParameterSetName, Sheet, SheetIndex)) {
            if (!ShouldProcess(sheet.Name, "Update Excel comments")) {
                continue;
            }

            updated += Run == null || Run.Length == 0
                ? sheet.UpdateComments(filter, Text!, Author, Initials)
                : sheet.UpdateCommentsRichText(filter, ExcelRichTextRunService.ToRuns(Run), Author, Initials);
        }

        if (updated > 0) {
            workbook.SaveIfOwned();
        }

        if (PassThru.IsPresent) {
            WriteObject(updated);
        }
    }

    private ExcelCommentFilter CreateRequiredFilter() {
        if (!string.IsNullOrWhiteSpace(Address) && !string.IsNullOrWhiteSpace(Range)) {
            throw new PSArgumentException("Specify either -Address or -Range, not both.");
        }

        bool hasFilter = !string.IsNullOrWhiteSpace(Address)
            || !string.IsNullOrWhiteSpace(Range)
            || !string.IsNullOrWhiteSpace(MatchAuthor)
            || !string.IsNullOrWhiteSpace(TextContains);
        if (!hasFilter && !All.IsPresent) {
            throw new PSArgumentException("Specify a comment filter or use -All.");
        }

        return new ExcelCommentFilter {
            A1Range = !string.IsNullOrWhiteSpace(Address) ? Address : Range,
            Author = MatchAuthor,
            TextContains = TextContains
        };
    }
}