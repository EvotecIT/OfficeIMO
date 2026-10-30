using System.IO;
using System.Management.Automation;
using System.Net;
using System.Security;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Opens a Word document for explicit editing, or saves and closes a completed content block.</summary>
/// <para>Opening preserves the document's theme and encryption. Without a block the document remains live. A failed block leaves the source unchanged; read-only input cannot run an authoring block.</para>
/// <example><summary>Open a literal file path from a file pipeline.</summary>
/// <code>$document = Get-Item './Review[final].docx' | Open-OfficeWord
/// $document | Close-OfficeWord</code></example>
/// <example><summary>Edit and save an existing document atomically.</summary>
/// <code>Open-OfficeWord -Path ./Report.docx -PassThru {
///     WordParagraph 'Reviewed and approved'
/// }</code></example>
[Cmdlet(VerbsCommon.Open, "OfficeWord", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(WordDocument), typeof(FileInfo))]
public sealed class OpenOfficeWordCommand : OfficeLifecycleCmdlet<WordDocument> {
    /// <summary>Word document path with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards]
    [ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;

    /// <summary>Word document path without wildcard expansion; accepts FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath")]
    [ValidateNotNullOrEmpty]
    public string LiteralPath { get; set; } = string.Empty;

    /// <summary>Return a read-only live document. Cannot be combined with Content.</summary>
    [Parameter]
    public SwitchParameter ReadOnly { get; set; }

    /// <summary>Source password as a SecureString, credential, or warned plaintext value.</summary>
    [Parameter]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ReadOnly && Content != null) {
            InvalidArgument(new PSArgumentException("ReadOnly cannot be combined with an authoring content block."), ReadOnly);
            return;
        }
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string input in OfficePath.ResolveInput(SessionState, literal ? LiteralPath : Path, literal)) {
            RunDocument(input, isNew: false, force: false,
                path => WordDocumentService.LoadDocument(path!, ReadOnly.IsPresent, autoSave: false,
                    Password == null ? null : new NetworkCredential(string.Empty, Password).Password),
                (document, path, allowLoss) => {
                    document.Save(path, WordDocumentService.NativeSaveOptions(allowLoss, null, passwordBound: false));
                    return null;
                });
        }
    }

    /// <inheritdoc />
    protected override void InvokeContent(WordDocument document) => WordDocumentService.InvokeDsl(document, Content!);
}
