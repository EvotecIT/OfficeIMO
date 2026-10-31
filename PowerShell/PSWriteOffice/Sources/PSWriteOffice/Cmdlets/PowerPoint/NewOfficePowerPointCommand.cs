using System.IO;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Creates a themed presentation, returning a live document or saving a completed content block.</summary>
/// <para>Without a content block no file is written, including on disposal. With a block, publication is atomic and an existing file requires Force. A block failure leaves the destination unchanged.</para>
/// <example>
///   <summary>Create a deck and return its saved file.</summary>
///   <code>New-OfficePowerPoint -Path ./Status.pptx -Theme Modern -PassThru {
///     PowerPointSlide -Title 'Status update'
/// }</code>
/// </example>
/// <example>
///   <summary>Create a live document for explicit editing and saving.</summary>
///   <code>$deck = New-OfficePowerPoint -Theme Boardroom
/// PowerPointSlide -Document $deck -Title 'Service review'
/// Save-OfficePowerPoint -Document $deck -Path ./Review.pptx
/// Close-OfficePowerPoint -Document $deck</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePowerPoint", SupportsShouldProcess = true)]
[OutputType(typeof(PowerPointDocument), typeof(FileInfo))]
public sealed class NewOfficePowerPointCommand : OfficeLifecycleCmdlet<PowerPointDocument> {
    /// <summary>Optional associated destination; required with a content block.</summary>
    [Parameter(Position = 0)]
    [ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>Allow the new document to replace an existing destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Built-in theme, theme object or a Word, Excel or PowerPoint document/template whose visual theme is read. Defaults to Modern.</summary>
    [Parameter]
    [OfficeThemeTransformation]
    public OfficeThemeInput? Theme { get; set; }

    /// <summary>Existing presentation/template copied as starting content. Theme applies visual styling separately.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Template { get; set; }

    /// <summary>Password for the saved file; plaintext input warns, and credentials contribute their password.</summary>
    [Parameter]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Content == null && MyInvocation.BoundParameters.ContainsKey(nameof(Password))) {
            InvalidArgument(new PSArgumentException("Password requires a content block. Supply it to Save-OfficePowerPoint when saving a live document."), Password);
            return;
        }
        var options = new PowerPointCreateOptions {
            PersistenceMode = DocumentPersistenceMode.Explicit,
            Theme = Theme?.Resolve(OfficeTheme.FromDocument),
            Template = Template == null ? null : ResolveOutputPath(Template)
        };
        RunDocument(Path, isNew: true, force: Force.IsPresent,
            path => path == null ? PowerPointDocument.Create(options) : PowerPointDocument.Create(path, options),
            (document, path, allowLoss) => {
                document.Save(path, PowerPointDocumentService.NativeSaveOptions(allowLoss, Password,
                    MyInvocation.BoundParameters.ContainsKey(nameof(Password))));
                return null;
            });
    }

    /// <inheritdoc />
    protected override void InvokeContent(PowerPointDocument document) =>
        PowerPointDocumentService.InvokeContent(document, Content!);
}
