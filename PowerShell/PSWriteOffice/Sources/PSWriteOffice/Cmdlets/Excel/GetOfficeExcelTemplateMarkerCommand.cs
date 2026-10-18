using System;
using System.Collections;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Lists Excel template markers such as {{Name}} and optionally shows whether supplied values bind to them.</summary>
/// <example>
///   <summary>Inspect template markers before filling a workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeExcelTemplateMarker -Path .\Invoice.xlsx -Sheet Invoice -Value @{ Number = 'INV-001'; Total = 123.45 }</code>
///   <para>Returns one object per marker with address, format, and binding metadata.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelTemplateMarker", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelTemplateMarkers")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelTemplateMarkerCommand : PSCmdlet {
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

    /// <summary>Optional marker values used to report which markers are bound and which are still missing.</summary>
    [Parameter]
    [Alias("Values")]
    public Hashtable? Value { get; set; }

    /// <summary>Only returns markers that are not supplied by -Value.</summary>
    [Parameter]
    public SwitchParameter MissingOnly { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: true);
        var path = string.Equals(ParameterSetName, ParameterSetPath, StringComparison.OrdinalIgnoreCase)
            ? Path
            : null;
        var values = Value == null
            ? null
            : ExcelTemplateValueService.ConvertValues(Value);

        foreach (var sheet in ExcelWorkbookCommandService.ResolveSheets(this, workbook.Document, ParameterSetName, Sheet, SheetIndex)) {
            var inspection = values == null
                ? sheet.InspectTemplate()
                : sheet.InspectTemplate(values);

            foreach (var marker in inspection.Markers) {
                if (MissingOnly.IsPresent && marker.IsBound == true) {
                    continue;
                }

                WriteObject(ExcelTemplateValueService.CreateMarkerRecord(marker, path));
            }
        }
    }
}