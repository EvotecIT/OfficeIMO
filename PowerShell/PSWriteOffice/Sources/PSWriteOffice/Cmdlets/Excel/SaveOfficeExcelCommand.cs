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
/// <example>
/// <summary>Save a live workbook</summary>
/// <code>$book = Open-OfficeExcel -Path .\Reports\Inventory.xlsx
/// Save-OfficeExcel -Document $book
/// Close-OfficeExcel -Document $book</code>
/// <para>Saves a workbook opened without a content block back to its own path and closes it.</para>
/// </example>
/// <example>
/// <summary>Save a validated copy under a new name</summary>
/// <code>$book = New-OfficeExcel
/// Add-OfficeExcelSheet -Document $book -Name 'Changes' {
///     ExcelTable (Import-OfficeCsv -Path .\Data\Changes-Week43.csv)
/// }
/// $book |
///     Save-OfficeExcel -Path .\Reports\Changes-Week43.xlsx -SaveOptions @{ ValidateOpenXml = $true } -PassThru
/// Close-OfficeExcel -Document $book</code>
/// <para>Writes this week's change log to a new file with Open XML validation, returns the saved file, and closes the live workbook explicitly.</para>
/// </example>
/// <example>
/// <summary>Save with a password</summary>
/// <code>$password = Read-Host -AsSecureString -Prompt 'Workbook password'
/// $book = Open-OfficeExcel -Path .\Reports\Salaries-2027.xlsx
/// Save-OfficeExcel -Document $book -Path .\Secure\Salaries-2027.xlsx -Password $password
/// Close-OfficeExcel -Document $book</code>
/// <para>Encrypts the saved file with a password read at the prompt, so it never appears in the script or history.</para>
/// </example>
[Cmdlet(VerbsData.Save, "OfficeExcel", SupportsShouldProcess = true, DefaultParameterSetName = "Default")]
[OutputType(typeof(FileInfo))]
public sealed class SaveOfficeExcelCommand : OfficeCmdlet {
    /// <summary>Live workbook to save.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [ValidateNotNull]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Optional new destination; omission uses the associated path.</summary>
    [Parameter(ParameterSetName = "Default")]
    [ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>Override the output password; omission retains loaded encryption.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>Native save settings, or a hashtable using the save-option factory's names.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeExcelSaveOptionsTransformation, ValidateNotNull]
    public ExcelSaveOptions? SaveOptions { get; set; }

    /// <summary>Explicitly accept reported native-save reductions; overrides SaveOptions when supplied.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowLoss { get; set; }

    /// <summary>Open the saved file in its registered default application.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Show { get; set; }

    /// <summary>Return the saved FileInfo.</summary>
    [Parameter(ParameterSetName = "Default")]
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
