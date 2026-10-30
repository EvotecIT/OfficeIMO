using System.IO;
using System.Management.Automation;
using System.Net;
using System.Security;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Opens a presentation for explicit editing, or atomically saves a completed content block.</summary>
/// <para>Opening preserves the document's theme and encryption. Without a block, the returned document stays open until Close-OfficePowerPoint. A failed block leaves the source unchanged.</para>
/// <example>
///   <summary>Open a deck without interpreting brackets as wildcards.</summary>
///   <code>$deck = Open-OfficePowerPoint -LiteralPath './Review[final].pptx'
/// Save-OfficePowerPoint -Document $deck -Path ./Review.pptx
/// Close-OfficePowerPoint -Document $deck</code>
/// </example>
[Cmdlet(VerbsCommon.Open, "OfficePowerPoint", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(PowerPointPresentation), typeof(FileInfo))]
public sealed class OpenOfficePowerPointCommand : OfficeLifecycleCmdlet<PowerPointPresentation> {
    /// <summary>Presentation paths with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards]
    [ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;

    /// <summary>Presentation path without wildcard expansion; accepts FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath")]
    [ValidateNotNullOrEmpty]
    public string LiteralPath { get; set; } = string.Empty;

    /// <summary>Source password as a SecureString, credential, or warned plaintext value.</summary>
    [Parameter]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string input in OfficePath.ResolveInput(SessionState, literal ? LiteralPath : Path, literal)) {
            RunDocument(input, isNew: false, force: false,
                path => PowerPointDocumentService.LoadPresentation(path!, Password == null ? null : new NetworkCredential(string.Empty, Password).Password),
                (document, path, allowLoss) => {
                    // The source password belongs to Load; omitting a save override
                    // keeps the native document's existing encryption policy.
                    document.Save(path, PowerPointDocumentService.NativeSaveOptions(allowLoss, null, passwordBound: false));
                    return null;
                });
        }
    }

    /// <inheritdoc />
    protected override void InvokeContent(PowerPointPresentation document) =>
        PowerPointDocumentService.InvokeContent(document, Content!);
}
