using System.Management.Automation;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Creates Excel/OpenDocument conversion settings.</summary>
/// <example>
///   <summary>Convert a bounded worksheet area with basic styles.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeExcelOpenDocumentOptions -IncludeBasicStyles -MaximumRows 10000 -MaximumColumns 100
/// ConvertTo-OfficeOpenDocument -Path .\Data.xlsx -OutputPath .\Data.ods -ExcelOptions $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeExcelOpenDocumentOptions")]
[OutputType(typeof(ExcelOpenDocumentConversionOptions))]
public sealed class NewOfficeExcelOpenDocumentOptionsCommand : PSCmdlet {
    /// <summary>Whether conversion loss is reported or rejected.</summary>
    [Parameter] public OdfConversionLossPolicy? LossPolicy { get; set; }
    /// <summary>Copy common font, fill, and number-format styles.</summary>
    [Parameter] public SwitchParameter IncludeBasicStyles { get; set; }
    /// <summary>Maximum cells materialized during conversion.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long? MaximumExpandedCells { get; set; }
    /// <summary>Maximum spreadsheet rows.</summary>
    [Parameter] [ValidateRange(1, 1048576)] public int? MaximumRows { get; set; }
    /// <summary>Maximum spreadsheet columns.</summary>
    [Parameter] [ValidateRange(1, 16384)] public int? MaximumColumns { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new ExcelOpenDocumentConversionOptions();
        if (LossPolicy.HasValue) options.LossPolicy = LossPolicy.Value;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(IncludeBasicStyles))) options.IncludeBasicStyles = IncludeBasicStyles.IsPresent;
        if (MaximumExpandedCells.HasValue) options.MaximumExpandedCells = MaximumExpandedCells.Value;
        if (MaximumRows.HasValue) options.MaximumRows = MaximumRows.Value;
        if (MaximumColumns.HasValue) options.MaximumColumns = MaximumColumns.Value;
        WriteObject(options);
    }
}
