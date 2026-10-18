using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Saves an OfficeIMO.Pdf document.</summary>
/// <remarks>
/// Use this command when a PDF is built in memory and saved later, or when a pipeline should continue with the saved file.
/// The document is saved through the normal OfficeIMO.Pdf save path.
/// </remarks>
/// <example>
///   <summary>Build a PDF in memory and save it later.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$pdf = New-OfficePdf { PdfHeading 'Queued report'; PdfParagraph 'Generated in memory.' }
/// $pdf | Save-OfficePdf -Path .\QueuedReport.pdf</code>
///   <para>Creates a PDF document object first, then saves it to disk.</para>
/// </example>
[Cmdlet(VerbsData.Save, "OfficePdf", SupportsShouldProcess = true)]
[OutputType(typeof(PdfDocument))]
public sealed class SaveOfficePdfCommand : PSCmdlet {
    /// <summary>PDF document to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Destination PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open the PDF after saving.</summary>
    [Parameter]
    [Alias("Show")]
    public SwitchParameter Open { get; set; }

    /// <summary>Emit the document for further processing.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Password required to open the generated PDF.</summary>
    [Parameter]
    [Alias("UserPassword")]
    public string? Password { get; set; }

    /// <summary>Optional owner password for the generated encrypted PDF.</summary>
    [Parameter]
    public string? OwnerPassword { get; set; }

    /// <summary>Raw PDF Standard security permission bit mask. Defaults to allowing all standard operations.</summary>
    [Parameter]
    [Alias("Permissions")]
    public int? Permission { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var fullPath = PdfCommandUtilities.ResolvePath(this, Path);
        if (!PdfCommandUtilities.ShouldWrite(this, fullPath, "Save PDF")) {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(fullPath);
        var document = PdfCommandUtilities.ApplyEncryption(Document, Password, OwnerPassword, Permission);
        document.Save(fullPath).RequireSuccess();

        if (Open.IsPresent) {
            FileOpenService.Open(fullPath);
        }

        if (PassThru.IsPresent) {
            WriteObject(document);
        }
    }
}