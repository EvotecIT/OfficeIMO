using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Saves an Excel workbook without disposing it.</summary>
/// <example>
///   <summary>Save a workbook in-place.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$workbook | Save-OfficeExcel</code>
///   <para>Writes pending changes to disk and keeps the workbook open.</para>
/// </example>
[Cmdlet(VerbsData.Save, "OfficeExcel", SupportsShouldProcess = true)]
[OutputType(typeof(ExcelDocument))]
public sealed class SaveOfficeExcelCommand : PSCmdlet {
    /// <summary>Workbook to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Optional save-as path.</summary>
    [Parameter]
    public string? Path { get; set; }

    /// <summary>Open the workbook after saving.</summary>
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

    /// <summary>Emit the workbook for further processing.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Document == null) {
            return;
        }

        var associatedPath = ExcelDocumentService.GetAssociatedPath(Document);
        if (string.IsNullOrWhiteSpace(Path) && string.IsNullOrWhiteSpace(associatedPath)) {
            throw new PSInvalidOperationException("No file path provided. Use -Path or open the workbook from disk.");
        }

        var saveOptions = ExcelDocumentService.CreateSaveOptions(
            SafePreflight.IsPresent,
            SafeRepairDefinedNames.IsPresent,
            ValidateOpenXml.IsPresent,
            DisableFastPackageWriter.IsPresent,
            EvaluateFormulas.IsPresent,
            ClearCachedFormulaResults.IsPresent,
            MarkFormulasDirty.IsPresent,
            ForceFullCalculationOnOpen.IsPresent);

        string savedPath;
        if (!string.IsNullOrWhiteSpace(Path)) {
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            if (!PdfCommandUtilities.ShouldWrite(this, resolvedPath, "Save Excel workbook")) {
                return;
            }

            if (string.IsNullOrEmpty(Password) &&
                ExcelDocumentService.IsEncryptedSource(Document) &&
                string.Equals(System.IO.Path.GetFullPath(resolvedPath), System.IO.Path.GetFullPath(associatedPath!), StringComparison.OrdinalIgnoreCase)) {
                throw new PSInvalidOperationException("Provide -Password when saving a workbook loaded from an encrypted package.");
            }

            ExcelDateSystemService.ApplyIfSpecified(Document, DateSystem, nameof(DateSystem));
            if (!string.IsNullOrEmpty(Password)) {
                OfficeEncryptedPackageService.SaveExcel(Document, resolvedPath, Password!, false, saveOptions);
            } else {
                Document.Save(resolvedPath, saveOptions);
            }
            savedPath = resolvedPath;
        } else {
            if (!PdfCommandUtilities.ShouldWrite(this, associatedPath!, "Save Excel workbook")) {
                return;
            }

            ExcelDateSystemService.ApplyIfSpecified(Document, DateSystem, nameof(DateSystem));
            if (!string.IsNullOrEmpty(Password)) {
                OfficeEncryptedPackageService.SaveExcel(Document, associatedPath!, Password!, false, saveOptions);
            } else {
                if (ExcelDocumentService.IsEncryptedSource(Document)) {
                    throw new PSInvalidOperationException("Provide -Password when saving a workbook loaded from an encrypted package.");
                }

                Document.Save(associatedPath!, saveOptions);
            }
            savedPath = associatedPath!;
        }

        ExcelDocumentService.UpdateSaveAssociation(Document, savedPath, !string.IsNullOrEmpty(Password));
        if (Open.IsPresent) {
            FileOpenService.Open(savedPath);
        }

        if (PassThru.IsPresent) {
            WriteObject(Document);
        }
    }

}