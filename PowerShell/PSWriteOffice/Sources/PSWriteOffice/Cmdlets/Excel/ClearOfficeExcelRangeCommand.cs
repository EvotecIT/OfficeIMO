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

/// <summary>Clears values, formulas, styles and range metadata from a range, a header column or whole sheets.</summary>
/// <remarks>Native table schemas and their header labels remain in place while the selected data and metadata categories are cleared.</remarks>
/// <example>
/// <summary>Clear the values of a range</summary>
/// <code>Open-OfficeExcel -Path .\Templates\Weekly-Report.xlsx {
///     ExcelSheet 'Data' {
///         Clear-OfficeExcelRange A2:F200 -Include Values
///     }
/// }</code>
/// <para>Removes the values from A2:F200 and keeps formatting, validation and comments.</para>
/// </example>
/// <example>
/// <summary>Clear worksheet data and rules</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Weekly-Ops.xlsx {
///     ExcelSheet 'Data' {
///         $cleared = Clear-OfficeExcelRange -PassThru
///         Write-Verbose "Selected $cleared cells"
///     }
/// }</code>
/// <para>Clears the supported data and rule categories on Data, retains table headers and reports the selected cell count.</para>
/// </example>
[Cmdlet(VerbsCommon.Clear, "OfficeExcelRange", DefaultParameterSetName = "Range")]
public sealed class ClearOfficeExcelRangeCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector[]? Sheet { get; set; }

    /// <summary>A1 cell range on the selected worksheet.</summary>
    [Parameter(ParameterSetName = "Range", Position = 0)]
    public string? Range { get; set; }

    /// <summary>Cell data and metadata categories to clear. Omission clears all supported categories.</summary>
    [Parameter()]
    public ExcelClearOptions[]? Include { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Header text identifying the target column.</summary>
    [Parameter(ParameterSetName = "Header", Mandatory = true)]
    public string Header { get; set; } = null!;

    /// <summary>Native Excel table used to resolve the source range or header.</summary>
    [Parameter(ParameterSetName = "Header")]
    public string? Table { get; set; }

    /// <summary>One-based row containing the column headers.</summary>
    [Parameter(ParameterSetName = "Header")]
    public int? HeaderRow { get; set; }

    /// <summary>Includes the header cell in the resolved column range.</summary>
    [Parameter(ParameterSetName = "Header")]
    public SwitchParameter IncludeHeader { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ExcelCommandTargets.Document(Document);
        ExcelClearOptions options = Include == null ? ExcelClearOptions.All : Include.Aggregate(ExcelClearOptions.None, (left, right) => left | right);
        int count = 0;
        foreach (var sheet in ExcelSheetResolver.ResolveMany(document, Sheet, Document == null ? ExcelDslContext.Current?.CurrentSheet : null)) {
            string target = ExcelTargetRangeResolver.ResolveOptional(sheet, Range, Header, Table, HeaderRow ?? 0, IncludeHeader) ?? sheet.UsedRange.Address;
            var (first, left, last, right) = ExcelA1.ParseRange(target.Contains(":") ? target : target + ":" + target);
            count += checked((last - first + 1) * (right - left + 1));
            sheet.Range(target).Clear(options);
        }
        if (PassThru) WriteObject(count);
    }

}
