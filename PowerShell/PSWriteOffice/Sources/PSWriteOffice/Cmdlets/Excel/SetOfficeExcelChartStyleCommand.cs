using System;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Applies a built-in style and color preset to an Excel chart.</summary>
/// <example>
///   <summary>Apply the default OfficeIMO preset.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$chart | Set-OfficeExcelChartStyle</code>
///   <para>Applies the default chart style and returns the chart for chaining.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChartStyle")]
[OutputType(typeof(ExcelChart))]
public sealed class SetOfficeExcelChartStyleCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Chart to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ExcelChart Chart { get; set; } = null!;

    /// <summary>Chart style identifier.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int StyleId { get; set; } = 251;

    /// <summary>Chart color style identifier.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int ColorStyleId { get; set; } = 10;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            Chart.ApplyStylePreset(StyleId, ColorStyleId);
            WritePassThru(Chart);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "ExcelChartStyleFailed", ErrorCategory.InvalidOperation, Chart));
        }
    }
}
