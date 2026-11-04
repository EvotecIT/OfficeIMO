using System;
using System.IO;
using System.Management.Automation;
using System.Security;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Saves a live Excel workbook without closing it.</summary>
/// <para>Omitted Password preserves loaded encryption; an explicit null removes it. SaveOptions accepts native options or a hashtable. PassThru returns the saved FileInfo. The workbook must use explicit persistence.</para>
/// <example><summary>Save pending edits and then close the workbook.</summary>
/// <code>Save-OfficeExcel -Document $workbook -Path ./Reviewed.xlsx -SaveOptions @{ Preflight = $true; ValidateOpenXml = $true } -PassThru
/// Close-OfficeExcel -Document $workbook</code></example>
[Cmdlet(VerbsData.Save, "OfficeExcel", SupportsShouldProcess = true, DefaultParameterSetName = "Default")]
[OutputType(typeof(FileInfo))]
public sealed class SaveOfficeExcelCommand : OfficeCmdlet {
    /// <summary>Live workbook to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0), ValidateNotNull]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Optional new destination; omission uses the associated path.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>Override the output password; omission retains loaded encryption.</summary>
    [Parameter, OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>Native save settings, or a hashtable using the save-option factory's names.</summary>
    [Parameter, OfficeExcelSaveOptionsTransformation, ValidateNotNull]
    public ExcelSaveOptions? SaveOptions { get; set; }

    /// <summary>Explicitly accept reported native-save reductions; overrides SaveOptions when supplied.</summary>
    [Parameter]
    public SwitchParameter AllowLoss { get; set; }

    /// <summary>Open the saved file in its registered default application.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Return the saved FileInfo.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        ExcelDocumentService.EnsureExplicitPersistence(Document);
        string? destination = Path ?? ExcelDocumentService.GetAssociatedPath(Document);
        if (string.IsNullOrWhiteSpace(destination)) {
            InvalidArgument(new PSArgumentException("Supply Path when the workbook has no associated destination."), Document);
            return;
        }
        destination = ResolveOutputPath(destination!);
        if (!ShouldProcess(destination, "Save Excel workbook")) return;
        try {
            var options = ExcelDocumentService.NativeSaveOptions(SaveOptions,
                MyInvocation.BoundParameters.ContainsKey(nameof(AllowLoss)) ? AllowLoss.IsPresent : (bool?)null,
                Password, MyInvocation.BoundParameters.ContainsKey(nameof(Password)));
            Document.Save(destination, options);
            ExcelDocumentService.UpdateSaveAssociation(Document, destination);
            if (Show) OfficeShow.Open(destination);
            if (PassThru) WriteObject(new FileInfo(destination));
        } catch (PipelineStoppedException) { throw; }
        catch (Exception error) { ItemError(error, "native-save-failed", destination); }
    }
}
