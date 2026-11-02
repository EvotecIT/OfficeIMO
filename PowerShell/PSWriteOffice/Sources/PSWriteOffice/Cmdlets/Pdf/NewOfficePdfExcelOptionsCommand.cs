using OfficeIMO.Pdf.Excel;
using System.Globalization;
using System.Management.Automation;
using OfficeIMO.Excel;
using OfficeIMO;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates discoverable PDF-table-to-Excel reconstruction settings.</summary>
/// <example><summary>Import PDF tables using British number and date parsing.</summary>
/// <code>$options = New-OfficePdfExcelOptions -Culture en-GB
/// Convert-OfficeDocument -Path ./Tables.pdf -To Excel -OutputPath ./Tables.xlsx -Options $options</code></example>
/// <example><summary>Keep values as text and disable fitted widths.</summary>
/// <code>$options = New-OfficePdfExcelOptions -NoNumericConversion -NoDateTimeConversion -NoAutoFit
/// Convert-OfficeDocument -Path ./Tables.pdf -To Excel -OutputPath ./Text.xlsx -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficePdfExcelOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(ExcelPdfImportOptions))]
public sealed class NewOfficePdfExcelOptionsCommand : OfficeCmdlet {
    /// <summary>Maximum body rows imported per detected table; zero means unlimited.</summary>
    [Parameter(ParameterSetName = "Default")] [ValidateRange(0, int.MaxValue)] public int? MaxRows { get; set; }
    /// <summary>Prefix for generated worksheet names.</summary>
    [Parameter] public string? SheetNamePrefix { get; set; }
    /// <summary>Prefix for generated Excel table names.</summary>
    [Parameter] public string? TableNamePrefix { get; set; }
    /// <summary>Excel table style.</summary>
    [Parameter] public ExcelTableStyle? TableStyle { get; set; }
    /// <summary>Exclude table-scoped AutoFilters.</summary>
    [Parameter] public SwitchParameter NoAutoFilter { get; set; }
    /// <summary>Keep native widths without fitting worksheet columns.</summary>
    [Parameter] public SwitchParameter NoAutoFit { get; set; }
    /// <summary>Keep numeric-looking values as text.</summary>
    [Parameter] public SwitchParameter NoNumericConversion { get; set; }
    /// <summary>Keep boolean-looking values as text.</summary>
    [Parameter] public SwitchParameter NoBooleanConversion { get; set; }
    /// <summary>Keep date-looking values as text.</summary>
    [Parameter] public SwitchParameter NoDateTimeConversion { get; set; }
    /// <summary>Keep percentage values as text.</summary>
    [Parameter] public SwitchParameter NoPercentageConversion { get; set; }
    /// <summary>Culture used for numeric parsing; defaults to the current culture.</summary>
    [Parameter, ValidateNotNull] public CultureInfo Culture { get; set; } = CultureInfo.CurrentCulture;
    /// <summary>Keep table segments on separate worksheets instead of merging across pages.</summary>
    [Parameter] public SwitchParameter NoPageContinuationMerge { get; set; }
    /// <summary>Retain repeated body header rows in merged table segments.</summary>
    [Parameter] public SwitchParameter KeepRepeatedHeaderRows { get; set; }
    /// <summary>Maximum table segments merged into one table.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaximumContinuationSegments { get; set; }
    /// <summary>Physical geometry tolerance for page continuations; bare numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength? ContinuationGeometryTolerance { get; set; }
    /// <summary>Worksheet name used when no tables are detected.</summary>
    [Parameter] public string? EmptyWorkbookSheetName { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new ExcelPdfImportOptions();
        Apply(nameof(NoAutoFilter), value => options.IncludeAutoFilter = !value);
        Apply(nameof(NoAutoFit), value => options.AutoFitColumns = !value);
        Apply(nameof(NoNumericConversion), value => options.ConvertNumericColumns = !value);
        Apply(nameof(NoBooleanConversion), value => options.ConvertBooleanColumns = !value);
        Apply(nameof(NoDateTimeConversion), value => options.ConvertDateTimeColumns = !value);
        Apply(nameof(NoPercentageConversion), value => options.ConvertPercentageColumns = !value);
        Apply(nameof(NoPageContinuationMerge), value => options.MergePageContinuations = !value);
        Apply(nameof(KeepRepeatedHeaderRows), value => options.SuppressRepeatedBodyHeaderRows = !value);
        if (MaxRows.HasValue) options.MaxRows = MaxRows.Value;
        if (SheetNamePrefix != null) options.SheetNamePrefix = SheetNamePrefix;
        if (TableNamePrefix != null) options.TableNamePrefix = TableNamePrefix;
        if (TableStyle.HasValue) options.TableStyle = TableStyle.Value;
        options.NumericCulture = (CultureInfo)Culture.Clone();
        if (MaximumContinuationSegments.HasValue) options.MaximumContinuationSegments = MaximumContinuationSegments.Value;
        if (ContinuationGeometryTolerance.HasValue) {
            double points = ContinuationGeometryTolerance.Value.ToPoints();
            if (points < 0) throw new PSArgumentException("ContinuationGeometryTolerance cannot be negative.");
            options.ContinuationGeometryTolerancePoints = points;
        }
        if (EmptyWorkbookSheetName != null) options.EmptyWorkbookSheetName = EmptyWorkbookSheetName;
        WriteObject(options);
    }

    private void Apply(string name, System.Action<bool> setter) {
        if (BoundSwitch(name) is bool value) setter(value);
    }
}
