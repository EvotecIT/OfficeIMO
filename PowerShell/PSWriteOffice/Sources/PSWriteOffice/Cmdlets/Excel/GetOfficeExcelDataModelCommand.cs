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

/// <summary>Reports connection, query-table, data-model and external-link parts in a workbook.</summary>
/// <example>
/// <summary>Report data connections</summary>
/// <code>Get-OfficeExcelDataModel -Path .\Reports\Live-Sales.xlsx</code>
/// <para>Lists connections, query tables, data-model parts and external links in the workbook.</para>
/// </example>
/// <example>
/// <summary>Find workbooks with external data</summary>
/// <code>Get-ChildItem -Path \\files\Finance -Filter *.xlsx -Recurse |
///     Get-OfficeExcelDataModel |
///     Format-List</code>
/// <para>Reports the data model of every workbook on the finance share, to find files that depend on databases or other workbooks.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelDataModel", DefaultParameterSetName = "Document")]
public sealed class GetOfficeExcelDataModelCommand : PSCmdlet {
    /// <summary>Workbook paths resolved with filesystem wildcard support.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string[] Path { get; set; } = Array.Empty<string>();

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
            WriteObject(document.InspectDataModel());
        }
    }

}
