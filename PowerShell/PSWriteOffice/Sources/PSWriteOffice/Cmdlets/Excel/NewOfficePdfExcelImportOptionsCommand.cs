using System.Globalization;
using System.Management.Automation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates discoverable PDF-table-to-Excel reconstruction settings.</summary>
/// <example>
///   <summary>Import PDF tables with typed columns and filters.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficePdfExcelImportOptions -IncludeAutoFilter -AutoFitColumns -ConvertNumericColumns -ConvertDateTimeColumns
/// ConvertTo-OfficePdfExcel -Path .\Tables.pdf -OutputPath .\Tables.xlsx -Options $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePdfExcelImportOptions")]
[OutputType(typeof(PdfTablesToExcelOptions))]
public sealed class NewOfficePdfExcelImportOptionsCommand : PSCmdlet {
    /// <summary>Maximum body rows imported per detected table; zero means unlimited.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int? MaxRows { get; set; }
    /// <summary>Prefix for generated worksheet names.</summary>
    [Parameter] public string? SheetNamePrefix { get; set; }
    /// <summary>Prefix for generated Excel table names.</summary>
    [Parameter] public string? TableNamePrefix { get; set; }
    /// <summary>Excel table style.</summary>
    [Parameter] public ExcelTableStyle? TableStyle { get; set; }
    /// <summary>Add table-scoped AutoFilters.</summary>
    [Parameter] public SwitchParameter IncludeAutoFilter { get; set; }
    /// <summary>Auto-fit worksheet columns.</summary>
    [Parameter] public SwitchParameter AutoFitColumns { get; set; }
    /// <summary>Convert consistently numeric columns.</summary>
    [Parameter] public SwitchParameter ConvertNumericColumns { get; set; }
    /// <summary>Convert consistently boolean columns.</summary>
    [Parameter] public SwitchParameter ConvertBooleanColumns { get; set; }
    /// <summary>Convert unambiguous date columns.</summary>
    [Parameter] public SwitchParameter ConvertDateTimeColumns { get; set; }
    /// <summary>Convert percentage columns to fractional numbers.</summary>
    [Parameter] public SwitchParameter ConvertPercentageColumns { get; set; }
    /// <summary>Culture name used for numeric parsing, such as en-US.</summary>
    [Parameter] public string? NumericCulture { get; set; }
    /// <summary>Merge compatible table segments across pages.</summary>
    [Parameter] public SwitchParameter MergePageContinuations { get; set; }
    /// <summary>Suppress repeated body header rows in merged segments.</summary>
    [Parameter] public SwitchParameter SuppressRepeatedBodyHeaderRows { get; set; }
    /// <summary>Maximum table segments merged into one table.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaximumContinuationSegments { get; set; }
    /// <summary>Geometry tolerance in PDF points for page continuations.</summary>
    [Parameter] [ValidateRange(0d, double.MaxValue)] public double? ContinuationGeometryTolerancePoints { get; set; }
    /// <summary>Worksheet name used when no tables are detected.</summary>
    [Parameter] public string? EmptyWorkbookSheetName { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new PdfTablesToExcelOptions();
        Apply(nameof(IncludeAutoFilter), value => options.IncludeAutoFilter = value);
        Apply(nameof(AutoFitColumns), value => options.AutoFitColumns = value);
        Apply(nameof(ConvertNumericColumns), value => options.ConvertNumericColumns = value);
        Apply(nameof(ConvertBooleanColumns), value => options.ConvertBooleanColumns = value);
        Apply(nameof(ConvertDateTimeColumns), value => options.ConvertDateTimeColumns = value);
        Apply(nameof(ConvertPercentageColumns), value => options.ConvertPercentageColumns = value);
        Apply(nameof(MergePageContinuations), value => options.MergePageContinuations = value);
        Apply(nameof(SuppressRepeatedBodyHeaderRows), value => options.SuppressRepeatedBodyHeaderRows = value);
        if (MaxRows.HasValue) options.MaxRows = MaxRows.Value;
        if (SheetNamePrefix != null) options.SheetNamePrefix = SheetNamePrefix;
        if (TableNamePrefix != null) options.TableNamePrefix = TableNamePrefix;
        if (TableStyle.HasValue) options.TableStyle = TableStyle.Value;
        if (!string.IsNullOrWhiteSpace(NumericCulture)) options.NumericCulture = CultureInfo.GetCultureInfo(NumericCulture!);
        if (MaximumContinuationSegments.HasValue) options.MaximumContinuationSegments = MaximumContinuationSegments.Value;
        if (ContinuationGeometryTolerancePoints.HasValue) options.ContinuationGeometryTolerancePoints = ContinuationGeometryTolerancePoints.Value;
        if (EmptyWorkbookSheetName != null) options.EmptyWorkbookSheetName = EmptyWorkbookSheetName;
        WriteObject(options);
    }

    private void Apply(string name, System.Action<bool> setter) {
        if (!MyInvocation.BoundParameters.ContainsKey(name)) return;
        setter(((SwitchParameter)MyInvocation.BoundParameters[name]).IsPresent);
    }
}
