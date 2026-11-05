using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets notes and threaded comments that match a filter.</summary>
/// <example>
/// <summary>List all comments</summary>
/// <code>Get-OfficeExcelComment -Path .\Finance\Forecast-2027.xlsx</code>
/// <para>Shows every note and threaded comment in the workbook.</para>
/// </example>
/// <example>
/// <summary>Collect open review threads</summary>
/// <code>Get-OfficeExcelComment -Path .\Finance\Forecast-2027.xlsx -Kind Threaded -Author 'Anna Kowalska' |
///     Export-OfficeExcel -Path .\Reviews\Forecast-Comments.xlsx -Title 'Open review comments'</code>
/// <para>Finds the threaded comments left by the finance reviewer and exports them as a review list.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelComment", DefaultParameterSetName = "Document")]
public sealed class GetOfficeExcelCommentCommand : PSCmdlet {
    /// <summary>Workbook paths resolved with filesystem wildcard support.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector[]? Sheet { get; set; }

    /// <summary>A1 cell range on the selected worksheet.</summary>
    [Parameter()]
    public string? Range { get; set; }

    /// <summary>Comment author to match, ignoring case.</summary>
    [Parameter()]
    public string? Author { get; set; }

    /// <summary>Text fragment to match, ignoring case.</summary>
    [Parameter()]
    public string? TextContains { get; set; }

    /// <summary>Metadata kind to add.</summary>
    [Parameter()]
    public ExcelCommentKind Kind { get; set; } = ExcelCommentKind.All;

    /// <summary>Literal workbook paths, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Document", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var scope in ExcelCommandTargets.Workbooks(this, ParameterSetName, Path, LiteralPath, Document, readOnly: true)) {
            var document = scope.Document;
            var selected = new HashSet<string>(ExcelSheetResolver.ResolveMany(document, Sheet, Document == null && !scope.OwnsDocument ? ExcelDslContext.Current?.CurrentSheet : null).Select(s => s.Name), StringComparer.OrdinalIgnoreCase);
            var report = document.InspectComments();
            if (Kind != ExcelCommentKind.Threaded) foreach (var sheet in document.Sheets.Where(s => selected.Contains(s.Name)))
                foreach (var comment in sheet.FindComments(new ExcelCommentFilter { A1Range = Range, Author = Author, TextContains = TextContains })) WriteObject(ExcelCommentRecordService.CreateRecord(comment, sheet.Name, document.FilePath));
            if (Kind != ExcelCommentKind.Note) foreach (var comment in report.ThreadedComments) {
                if (!selected.Contains(comment.SheetName)) continue;
                if (Author != null && !string.Equals(comment.Author, Author, StringComparison.OrdinalIgnoreCase)) continue;
                if (TextContains != null && comment.Text.IndexOf(TextContains, StringComparison.OrdinalIgnoreCase) < 0) continue;
                if (Range != null) {
                    var (r1, c1, r2, c2) = ExcelA1.ParseRange(Range.Contains(":") ? Range : Range + ":" + Range);
                    var (r, c) = ExcelA1.ParseCellRef(comment.CellReference);
                    if (r < r1 || r > r2 || c < c1 || c > c2) continue;
                }
                WriteObject(comment);
            }
        }
    }

}
