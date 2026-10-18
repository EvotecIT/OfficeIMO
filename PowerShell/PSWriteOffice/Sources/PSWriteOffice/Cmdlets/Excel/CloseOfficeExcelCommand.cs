using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Closes an Excel workbook and optionally saves it.</summary>
/// <para>Convenience wrapper so scripts do not need to call <see cref="ExcelDocument.Save()"/> or <c>Dispose</c> directly.</para>
/// <example>
///   <summary>Save, validate, and close an open workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeExcel -Path .\report.xlsx {
///     Add-OfficeExcelSheet -Name Data {
///         Set-OfficeExcelRow -Row 1 -Values 'Region', 'Revenue'
///         Set-OfficeExcelRow -Row 2 -Values 'EMEA', 98000
///     }
/// }
/// $workbook = Get-OfficeExcel -Path .\report.xlsx
/// $workbook | Close-OfficeExcel -Save -Path .\report-final.xlsx -SafePreflight -ValidateOpenXml</code>
///   <para>Saves pending changes through OfficeIMO's normal save path, validates the package, and releases the workbook.</para>
/// </example>
[Cmdlet(VerbsCommon.Close, "OfficeExcel", SupportsShouldProcess = true)]
public sealed class CloseOfficeExcelCommand : PSCmdlet {
    /// <summary>Workbook to close.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Persist changes before closing.</summary>
    [Parameter]
    public SwitchParameter Save { get; set; }

    /// <summary>Optional output path when saving.</summary>
    [Parameter]
    public string? Path { get; set; }

    /// <summary>Open the workbook after saving. Requires -Save or -Path.</summary>
    [Parameter]
    [Alias("Show")]
    public SwitchParameter Open { get; set; }

    /// <summary>Password used to save the workbook as an encrypted package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>Run OfficeIMO worksheet preflight cleanup before saving.</summary>
    [Parameter]
    public SwitchParameter SafePreflight { get; set; }

    /// <summary>Repair common defined-name issues before saving.</summary>
    [Parameter]
    public SwitchParameter SafeRepairDefinedNames { get; set; }

    /// <summary>Validate the saved package with OpenXmlValidator and throw on errors.</summary>
    [Parameter]
    public SwitchParameter ValidateOpenXml { get; set; }

    /// <summary>Disable OfficeIMO fast package writers for this save.</summary>
    [Parameter]
    public SwitchParameter DisableFastPackageWriter { get; set; }

    /// <summary>Evaluate supported formulas and write cached values before saving.</summary>
    [Parameter]
    public SwitchParameter EvaluateFormulas { get; set; }

    /// <summary>Remove cached formula results before saving.</summary>
    [Parameter]
    public SwitchParameter ClearCachedFormulaResults { get; set; }

    /// <summary>Mark formula cells dirty before saving.</summary>
    [Parameter]
    public SwitchParameter MarkFormulasDirty { get; set; }

    /// <summary>Request a full workbook recalculation when opened in Excel-compatible applications.</summary>
    [Parameter]
    public SwitchParameter ForceFullCalculationOnOpen { get; set; }

    /// <summary>Workbook date system for Excel date serials.</summary>
    [Parameter]
    [ValidateSet("1900", "1904", "NineteenHundred", "NineteenFour")]
    public string? DateSystem { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Document == null) {
            return;
        }

        if (Open.IsPresent && !Save.IsPresent && string.IsNullOrWhiteSpace(Path)) {
            throw new PSArgumentException("Use -Save or -Path with -Open so the workbook is persisted before it is opened.", nameof(Open));
        }

        var shouldSave = Save.IsPresent || !string.IsNullOrWhiteSpace(Path);
        var action = shouldSave ? "Save and close" : "Close";
        if (!ShouldProcess("Excel workbook", action)) {
            return;
        }

        if (shouldSave) {
            ExcelDateSystemService.ApplyIfSpecified(Document, DateSystem, nameof(DateSystem));
            var resolvedPath = !string.IsNullOrWhiteSpace(Path)
                ? SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path)
                : ExcelDocumentService.GetAssociatedPath(Document);
            var saveOptions = ExcelDocumentService.CreateSaveOptions(
                SafePreflight.IsPresent,
                SafeRepairDefinedNames.IsPresent,
                ValidateOpenXml.IsPresent,
                DisableFastPackageWriter.IsPresent,
                EvaluateFormulas.IsPresent,
                ClearCachedFormulaResults.IsPresent,
                MarkFormulasDirty.IsPresent,
                ForceFullCalculationOnOpen.IsPresent);
            ExcelDocumentService.SaveDocument(Document, Open.IsPresent, resolvedPath, Password, saveOptions);
        } else {
            ExcelDocumentService.CloseDocument(Document);
        }
    }
}
