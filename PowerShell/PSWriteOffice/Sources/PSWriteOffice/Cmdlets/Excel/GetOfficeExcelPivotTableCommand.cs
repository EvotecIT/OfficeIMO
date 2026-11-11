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

/// <summary>Lists pivot tables in a workbook.</summary>
/// <example>
/// <summary>List pivot tables</summary>
/// <code>Get-OfficeExcelPivotTable -Path .\Reports\Sales-2026.xlsx</code>
/// <para>Shows every pivot table in the workbook with its location and source.</para>
/// </example>
/// <example>
/// <summary>Check where pivot tables read from</summary>
/// <code>Get-OfficeExcelPivotTable -Path .\Reports\Sales-2026.xlsx |
///     Select-Object Name, SheetName, Location, SourceSheet, SourceRange</code>
/// <para>Lists the source range of each pivot table, to confirm they still point at the current data after the sheet was rebuilt.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelPivotTable", DefaultParameterSetName = "Document")]
public sealed class GetOfficeExcelPivotTableCommand : PSCmdlet {
    /// <summary>Workbook paths resolved with filesystem wildcard support.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Optional pivot table name filter.</summary>
    [Parameter()]
    public string? Name { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    [Alias("WorksheetName")]
    public OfficeSheetSelector? Sheet { get; set; }

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
            foreach (var table in document.GetPivotTables()) {
                if (Name != null && !new WildcardPattern(Name, WildcardOptions.IgnoreCase).IsMatch(table.Name)) continue;
                if (Sheet != null && !string.Equals(Sheet.Resolve(document.Sheets, s => s.Name).Name, table.SheetName, StringComparison.OrdinalIgnoreCase)) continue;
                WriteObject(table);
            }
        }
    }

}
