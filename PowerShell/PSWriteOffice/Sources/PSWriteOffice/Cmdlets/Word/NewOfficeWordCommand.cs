using System.IO;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates a themed Word document, returning a live document or saving a completed content block.</summary>
/// <para>Without a block no file is written, including on disposal. A block saves atomically; an existing destination requires Force. Template copies content, while Theme reads only the source theme.</para>
/// <example><summary>Create and save a report.</summary>
/// <code>New-OfficeWord -Path ./Report.docx -Theme Modern -PassThru {
///     WordHeading 'Service review'
///     WordParagraph 'All services are available.'
/// }</code></example>
/// <example><summary>Create a live document for explicit editing and saving.</summary>
/// <code>$document = New-OfficeWord -Theme Boardroom
/// $document | Add-OfficeWordParagraph -Text 'Reviewed and approved'
/// Save-OfficeWord -Document $document -Path ./Review.docx
/// Close-OfficeWord -Document $document</code></example>
[Cmdlet(VerbsCommon.New, "OfficeWord", DefaultParameterSetName = "Live", SupportsShouldProcess = true)]
[OutputType(typeof(WordDocument), typeof(FileInfo))]
public sealed class NewOfficeWordCommand : OfficeNewLifecycleCmdlet<WordDocument> {
    /// <summary>Built-in theme, theme object or a document/template whose visual theme is read.</summary>
    [Parameter, OfficeThemeTransformation]
    public OfficeThemeInput? Theme { get; set; }

    /// <summary>Existing Word package copied as starting content. Its styling is preserved unless Theme is supplied.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Template { get; set; }

    /// <summary>Allow a completed content block to replace an existing destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Password for a completed content block's saved file; use Save-OfficeWord for a live document.</summary>
    [Parameter(ParameterSetName = "Content"), OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new WordCreateOptions {
            PersistenceMode = DocumentPersistenceMode.Explicit,
            Theme = Theme?.Resolve(OfficeTheme.FromDocument),
            Template = Template == null ? null : ResolveOutputPath(Template)
        };
        RunDocument(Path, isNew: true, force: Force.IsPresent,
            path => path == null ? WordDocument.Create(options) : WordDocument.Create(path, options),
            (document, path, allowLoss) => {
                document.Save(path, WordDocumentService.NativeSaveOptions(allowLoss, Password,
                    MyInvocation.BoundParameters.ContainsKey(nameof(Password))));
                return null;
            });
    }

    /// <inheritdoc />
    protected override void InvokeContent(WordDocument document) => WordDocumentService.InvokeDsl(document, Content!);
}
