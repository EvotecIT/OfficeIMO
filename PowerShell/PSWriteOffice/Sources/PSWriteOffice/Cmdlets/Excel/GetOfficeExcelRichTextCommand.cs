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

/// <summary>Gets the rich-text runs of a cell.</summary>
/// <example>
/// <summary>Read the text runs of a cell</summary>
/// <code>Get-OfficeExcelRichText -Path .\Reports\Migration.xlsx -Sheet Overview -Address B3</code>
/// <para>Returns each formatted run of B3 on the Overview sheet.</para>
/// </example>
/// <example>
/// <summary>Find the bold parts of a status cell</summary>
/// <code>$book = Open-OfficeExcel -Path .\Reports\Status.xlsx -ReadOnly
/// Get-OfficeExcelRichText -Document $book -Sheet Summary -Address B2 | Where-Object Bold
/// Close-OfficeExcel -Document $book</code>
/// <para>Reads the runs of a status cell from an open workbook and keeps only the bold text.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelRichText", DefaultParameterSetName = "Document")]
public sealed class GetOfficeExcelRichTextCommand : PSCmdlet {
    /// <summary>Workbook paths resolved with filesystem wildcard support.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    [Alias("WorksheetName")]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>A1 address of the target cell.</summary>
    [Parameter(Mandatory = true)]
    public string Address { get; set; } = null!;

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
            var current = Document == null && !scope.OwnsDocument ? ExcelDslContext.Current?.CurrentSheet : null;
            foreach (var sheet in ExcelSheetResolver.ResolveMany(document, Sheet == null ? null : new[] { Sheet }, current)) {
                var (row, column) = ExcelHostExtensions.ResolveCellAddress(null, null, Address);
                var runs = sheet.Cell(row, column).GetRichText();
                for (int i = 0; i < runs.Count; i++) WriteObject(ExcelRichTextRunService.CreateRecord(runs[i], i, Address, row, column, sheet.Name, document.FilePath));
            }
        }
    }

}
