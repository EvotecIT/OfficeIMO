using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates native import policies and typed worksheet selections for Merge-OfficeDocument.</summary>
/// <para>The first input or live Document remains the base. Sheet selects appended source worksheets;
/// sheet objects identify worksheets in saved source files and remain caller-owned.</para>
/// <example><summary>Append selected worksheets while retaining native package content.</summary>
/// <code>$options = New-OfficeExcelMergeOptions -Sheet 'Sales*', 3 -NamePrefix 'Imported ' -CopyMode Package
/// Merge-OfficeDocument -Path ./Cover.xlsx, ./North.xlsx, ./South.xlsx -OutputPath ./Combined.xlsx -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficeExcelMergeOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(ExcelMergeOptions))]
public sealed class NewOfficeExcelMergeOptionsCommand : OfficeCmdlet {
    /// <summary>Source worksheet names/wildcards, one-based positions or sheet objects from saved appended sources.</summary>
    [Parameter(ParameterSetName = "Default"), OfficeSheetTransformation]
    public OfficeSheetSelector[]? Sheet { get; set; }
    /// <summary>Prefix added to imported worksheet names.</summary>
    [Parameter]
    public string? NamePrefix { get; set; }
    /// <summary>Native imported-sheet name validation policy.</summary>
    [Parameter]
    public ExcelSheetNameValidationMode? ValidationMode { get; set; }
    /// <summary>Native value or package copy policy.</summary>
    [Parameter]
    public ExcelWorksheetCopyMode? CopyMode { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var native = new ExcelWorkbookMergeOptions { SheetNamePrefix = NamePrefix };
        if (ValidationMode.HasValue) native.SheetNameValidationMode = ValidationMode.Value;
        if (CopyMode.HasValue) native.CopyMode = CopyMode.Value;
        WriteObject(new ExcelMergeOptions(native, Sheet));
    }
}
