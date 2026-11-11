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

/// <summary>Gets built-in, application and custom document properties.</summary>
/// <example>
/// <summary>Show document properties</summary>
/// <code>Get-OfficeExcelDocumentProperty -Path .\Reports\Q3-Sales.xlsx</code>
/// <para>Lists the built-in, application and custom properties of a workbook.</para>
/// </example>
/// <example>
/// <summary>Report owners of shared workbooks</summary>
/// <code>Get-ChildItem -Path \\files\Reports -Filter *.xlsx |
///     Get-OfficeExcelDocumentProperty -Name Owner, Classification -Kind Custom |
///     Format-Table</code>
/// <para>Reads the custom Owner and Classification properties of every workbook on the share.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelDocumentProperty", DefaultParameterSetName = "Document")]
public sealed class GetOfficeExcelDocumentPropertyCommand : PSCmdlet {
    /// <summary>Workbook paths resolved with filesystem wildcard support.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Property name filter (wildcards supported).</summary>
    [Parameter()]
    public string[]? Name { get; set; }

    /// <summary>Metadata kind to add.</summary>
    [Parameter()]
    public ExcelDocumentPropertyKind[]? Kind { get; set; }

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
            var kinds = new HashSet<ExcelDocumentPropertyKind>(Kind ?? (ExcelDocumentPropertyKind[])Enum.GetValues(typeof(ExcelDocumentPropertyKind)));
            foreach (var property in ExcelDocumentPropertyService.GetProperties(scope.Document, kinds.Contains(ExcelDocumentPropertyKind.BuiltIn), kinds.Contains(ExcelDocumentPropertyKind.Application), kinds.Contains(ExcelDocumentPropertyKind.Custom))) {
                if (Name == null || Name.Any(name => new WildcardPattern(name, WildcardOptions.IgnoreCase).IsMatch(property.Name))) WriteObject(property);
            }
        }
    }

}
