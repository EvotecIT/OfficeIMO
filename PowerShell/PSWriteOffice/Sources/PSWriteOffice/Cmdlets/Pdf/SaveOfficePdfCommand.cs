using System;
using System.IO;
using System.Management.Automation;
using System.Security;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Saves a live PDF to its associated path or a new destination without closing it.</summary>
/// <para>Native serialization rejects content loss unless AllowLoss is supplied. Password and OwnerPassword accept shared secure input. Security options encrypt an unencrypted output; use Protect-OfficeDocument for security changes to an already protected file.</para>
/// <example><summary>Save to the associated destination and keep editing.</summary>
/// <code>$pdf = New-OfficePdf -Path ./Review.pdf
/// Add-OfficePdfParagraph -Document $pdf -Text 'Reviewed and approved'
/// $pdf | Save-OfficePdf -PassThru
/// Close-OfficePdf -Document $pdf</code></example>
/// <example><summary>Restrict printing and accessibility without requiring an opening password.</summary>
/// <code>$owner = Read-Host -AsSecureString
/// Save-OfficePdf -Document $pdf -Path ./Restricted.pdf -OwnerPassword $owner -Permission Print,Accessibility</code></example>
[Cmdlet(VerbsData.Save, "OfficePdf", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class SaveOfficePdfCommand : OfficeCmdlet {
    /// <summary>Live PDF to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true), ValidateNotNull]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Optional destination; omission uses the native document's associated path.</summary>
    [Parameter(Position = 0), ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>Open the saved file in its registered default application.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Return the saved FileInfo.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Explicitly accept fidelity loss from native PDF serialization.</summary>
    [Parameter]
    public SwitchParameter AllowLoss { get; set; }

    /// <summary>Opening password for an unencrypted output; accepts secure input or warned plaintext.</summary>
    [Parameter, OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>Owner password; without Password the PDF opens without a password and enforces permissions.</summary>
    [Parameter, OfficePasswordTransformation]
    public SecureString? OwnerPassword { get; set; }

    /// <summary>Operations allowed to readers, expressed as native PDF permission flags.</summary>
    [Parameter]
    public PdfStandardPermissions Permission { get; set; } = PdfStandardPermissions.All;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        string? destination = Path ?? Document.FilePath;
        if (string.IsNullOrWhiteSpace(destination)) {
            InvalidArgument(new PSArgumentException("Supply Path when the document has no associated destination."), Document);
            return;
        }
        destination = ResolveOutputPath(destination!);
        var options = new PdfSaveOptions {
            AllowLoss = AllowLoss.IsPresent,
            Encryption = PdfCommandUtilities.CreateEncryption(Password, OwnerPassword, Permission,
                MyInvocation.BoundParameters.ContainsKey(nameof(Permission)))
        };
        if (!ShouldProcess(destination, "Save PDF document")) return;
        try {
            Document.Save(destination, options);
            if (Show) OfficeShow.Open(destination);
            if (PassThru) WriteObject(new FileInfo(destination));
        } catch (PipelineStoppedException) { throw; }
        catch (Exception error) { ItemError(error, "native-save-failed", destination); }
    }
}
