using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Clears legacy worksheet comments (notes) that match a filter.</summary>
/// <example>
///   <summary>Clear comments containing review text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$removed = Clear-OfficeExcelComment -Path .\Report.xlsx -Sheet Data -TextContains review -Confirm:$false -PassThru
/// Get-OfficeExcelCommentAudit -Path .\Report.xlsx -IncludeComments |
///     Select-Object LegacyCommentCount, ThreadedCommentCount</code>
///   <para>Removes matching comments and saves the workbook.</para>
/// </example>
[Cmdlet(VerbsCommon.Clear, "OfficeExcelComment", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Medium, DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelCommentClear")]
[OutputType(typeof(int))]
public sealed class ClearOfficeExcelCommentCommand : PSCmdlet {
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

    /// <summary>Comment author to match, ignoring case.</summary>
    [Parameter]
    public string? Author { get; set; }

    /// <summary>Text fragment to match, ignoring case.</summary>
    [Parameter]
    public string? TextContains { get; set; }

    /// <summary>Allow clearing all comments on the selected worksheet(s) when no filter is supplied.</summary>
    [Parameter]
    public SwitchParameter All { get; set; }

    /// <summary>Returns the number of comments cleared.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var filter = CreateRequiredFilter();
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        var cleared = 0;
        var shouldSave = false;

        foreach (var sheet in ExcelWorkbookCommandService.ResolveSheets(this, workbook.Document, ParameterSetName, Sheet, SheetIndex)) {
            if (!ShouldProcess(sheet.Name, "Clear Excel comments")) {
                continue;
            }

            shouldSave = true;
            cleared += sheet.ClearComments(filter);
        }

        if (shouldSave) {
            workbook.SaveIfOwned();
        }

        if (PassThru.IsPresent) {
            WriteObject(cleared);
        }
    }

    private ExcelCommentFilter CreateRequiredFilter() {
        if (!string.IsNullOrWhiteSpace(Address) && !string.IsNullOrWhiteSpace(Range)) {
            throw new PSArgumentException("Specify either -Address or -Range, not both.");
        }

        bool hasFilter = !string.IsNullOrWhiteSpace(Address)
            || !string.IsNullOrWhiteSpace(Range)
            || !string.IsNullOrWhiteSpace(Author)
            || !string.IsNullOrWhiteSpace(TextContains);
        if (!hasFilter && !All.IsPresent) {
            throw new PSArgumentException("Specify a comment filter or use -All.");
        }

        return new ExcelCommentFilter {
            A1Range = !string.IsNullOrWhiteSpace(Address) ? Address : Range,
            Author = Author,
            TextContains = TextContains
        };
    }
}