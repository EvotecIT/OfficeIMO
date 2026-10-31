using System.IO;
using System.Management.Automation;
using System.Security;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Saves a live presentation without closing it.</summary>
/// <para>Uses the supplied path or the document's associated destination. Omitted Password preserves loaded encryption. PassThru returns the saved FileInfo.</para>
/// <example>
///   <summary>Save a live document to a new destination.</summary>
///   <code>$deck = New-OfficePowerPoint
/// Save-OfficePowerPoint -Document $deck -Path ./Review.pptx -PassThru
/// Close-OfficePowerPoint -Document $deck</code>
/// </example>
[Cmdlet(VerbsData.Save, "OfficePowerPoint", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class SaveOfficePowerPointCommand : OfficeDocumentCmdlet<PowerPointDocument> {
    /// <summary>Optional new destination; omission uses the associated path.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>Override the output password; omission preserves the document's loaded encryption.</summary>
    [Parameter]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>Explicitly accept reported native-save reductions.</summary>
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
        PowerPointDocument document = GetDocument();
        string? destination = Path ?? document.FilePath;
        if (string.IsNullOrWhiteSpace(destination)) {
            InvalidArgument(new System.ArgumentException("Supply -Path when the document has no associated destination."), document, "destination-required");
            return;
        }
        destination = ResolveOutputPath(destination!);
        if (!ShouldProcess(destination, "Save PowerPoint presentation")) return;
        try {
            document.Save(destination, PowerPointDocumentService.NativeSaveOptions(AllowLoss.IsPresent, Password,
                MyInvocation.BoundParameters.ContainsKey(nameof(Password))));
            if (Show) OfficeShow.Open(destination);
            if (PassThru) WriteObject(new FileInfo(destination));
        } catch (PipelineStoppedException) { throw; }
        catch (System.Exception error) { ItemError(error, "native-save-failed", destination); }
    }
}
