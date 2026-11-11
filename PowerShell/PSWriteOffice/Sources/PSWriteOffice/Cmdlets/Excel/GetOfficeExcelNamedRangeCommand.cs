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

/// <summary>Gets defined names from a workbook.</summary>
/// <example>
/// <summary>List defined names</summary>
/// <code>Get-OfficeExcelNamedRange -Path .\Finance\Pricing-Model.xlsx</code>
/// <para>Shows every defined name and what it refers to.</para>
/// </example>
/// <example>
/// <summary>Find broken names</summary>
/// <code>Get-OfficeExcelNamedRange -Path .\Finance\Budget-2027.xlsx |
///     Where-Object Reference -like '*#REF!*' |
///     Select-Object Name, SheetName, Reference</code>
/// <para>Lists the defined names that refer to #REF! after sheets were deleted, so they can be repaired or removed.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelNamedRange", DefaultParameterSetName = "Document")]
public sealed class GetOfficeExcelNamedRangeCommand : AsyncPSCmdlet {
    /// <summary>Workbook paths resolved with filesystem wildcard support.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Optional named range to retrieve.</summary>
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
        var scope = Sheet == null ? null : Sheet.Resolve(document.Sheets, s => s.Name);
        foreach (var entry in document.ListNamedRanges()) {
            if (!string.Equals(entry.SheetName, scope?.Name, StringComparison.OrdinalIgnoreCase)) continue;
            if (Name != null && !new WildcardPattern(Name, WildcardOptions.IgnoreCase).IsMatch(entry.Name)) continue;
            var item = PSObject.AsPSObject(entry);
            ExcelRange? range = document.GetNamedRange(entry.Name, scope);
            item.Properties.Add(new PSNoteProperty("Range", range?.Address ?? entry.Reference));
            item.Properties.Add(new PSNoteProperty("Scope", scope?.Name ?? "Workbook"));
            item.Properties.Add(new PSNoteProperty("Sheet", range?.Sheet.Name ?? scope?.Name));
            if (document.FilePath != null) item.Properties.Add(new PSNoteProperty("Path", document.FilePath));
            if (ParameterSetName == "Uri") item.Properties.Add(new PSNoteProperty("Uri", Uri));
            WriteObject(item);
        }
    }

}
