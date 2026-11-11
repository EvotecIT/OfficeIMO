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

/// <summary>Lists the tables defined in a workbook.</summary>
/// <example>
/// <summary>List the tables in a workbook</summary>
/// <code>Get-OfficeExcelTable -Path .\Reports\Q3-Sales.xlsx</code>
/// <para>Shows every table with its sheet, range and style.</para>
/// </example>
/// <example>
/// <summary>Inventory the tables across report files</summary>
/// <code>Get-ChildItem -Path .\Reports -Filter *.xlsx |
///     Get-OfficeExcelTable |
///     Select-Object SheetName, Name, Range, StyleName</code>
/// <para>Lists the tables of every workbook in the Reports folder with where they live, to plan which ones a dashboard should read.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelTable", DefaultParameterSetName = "Document")]
public sealed class GetOfficeExcelTableCommand : AsyncPSCmdlet {
    /// <summary>Workbook paths resolved with filesystem wildcard support.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Optional table name filter.</summary>
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

    /// <summary>Remote workbook URI to inspect.</summary>
    [Parameter(ParameterSetName = "Uri", Mandatory = true, Position = 0)]
    public Uri Uri { get; set; } = null!;

    /// <summary>Allow HTTP workbook downloads in addition to HTTPS.</summary>
    [Parameter(ParameterSetName = "Uri")]
    public SwitchParameter AllowHttp { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        if (ParameterSetName == "Uri") {
            using var document = await ExcelDocumentService.LoadDocumentAsync(Uri!, readOnly: true, allowHttp: AllowHttp, cancellationToken: CancelToken).ConfigureAwait(false);
            Query(document);
        } else foreach (var scope in ExcelCommandTargets.Workbooks(this, ParameterSetName, Path, LiteralPath, Document, readOnly: true)) Query(scope.Document);
    }
    private void Query(ExcelDocument document) {
        foreach (var table in document.GetTables()) {
            if (Name != null && !new WildcardPattern(Name, WildcardOptions.IgnoreCase).IsMatch(table.Name)) continue;
            if (Sheet != null && !string.Equals(Sheet.Resolve(document.Sheets, s => s.Name).Name, table.SheetName, StringComparison.OrdinalIgnoreCase)) continue;
            var record = PSObject.AsPSObject(table);
            if (document.FilePath != null) record.Properties.Add(new PSNoteProperty("Path", document.FilePath));
            if (ParameterSetName == "Uri") record.Properties.Add(new PSNoteProperty("Uri", Uri));
            WriteObject(record);
        }
    }

}
