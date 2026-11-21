using System.Management.Automation;
using OfficeIMO.Toml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Toml;

namespace PSWriteOffice.Cmdlets.Toml;

/// <summary>Opens a TOML file, stream or text as a lossless live document; a file content block saves atomically.</summary>
/// <para>Saving writes only what changed: comments, blank lines, quoting, key order, alignment and the spelling of
/// every other value stay byte for byte. Malformed input fails with an error ID built from the engine's code (for
/// example toml-Toml.Load.DuplicateKey), and the message gives the line and column.</para>
/// <example><summary>Change one value and keep the rest of the file as it was.</summary>
/// <code>Open-OfficeToml -Path ./pyproject.toml {
///     Set-OfficeTomlNode -Select 'project.version' -Value '1.2.0'
/// }</code></example>
/// <example><summary>Read an uploaded file with untrusted limits.</summary>
/// <code>$document = Open-OfficeToml -LiteralPath ./uploads/Cargo.toml -Untrusted
/// Get-OfficeTomlNode -Document $document -Select 'package.name'
/// Close-OfficeToml -Document $document</code></example>
[Cmdlet(VerbsCommon.Open, "OfficeToml", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(TomlDocument), typeof(FileInfo))]
public sealed class OpenOfficeTomlCommand : OfficeLifecycleCmdlet<TomlDocument> {
    /// <summary>An optional editing block; omission returns the live document.</summary>
    [Parameter(ParameterSetName = "Path", Position = 1)]
    [Parameter(ParameterSetName = "LiteralPath", Position = 1)]
    [ValidateNotNull]
    public override ScriptBlock? Content { get => base.Content; set => base.Content = value; }

    /// <summary>Emit the FileInfo after a completed editing block.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public override SwitchParameter PassThru { get => base.PassThru; set => base.PassThru = value; }

    /// <summary>Open the saved file after a completed editing block.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public override SwitchParameter Show { get => base.Show; set => base.Show = value; }

    /// <summary>TOML file with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;
    /// <summary>Literal TOML file, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string LiteralPath { get; set; } = string.Empty;
    /// <summary>TOML text to parse without a file.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>A readable stream, read from its current position; it stays open.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Stream"), ValidateNotNull]
    public Stream? Stream { get; set; }
    /// <summary>Native load options or a hashtable of their properties: MaxDepth, MaxNodes, MaxStringLength and MaxInputBytes.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(TomlLoadOptions)), ValidateNotNull]
    public TomlLoadOptions? Options { get; set; }
    /// <summary>Caps every limit at the engine's untrusted defaults, for files from outside.</summary>
    [Parameter]
    public SwitchParameter Untrusted { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        TomlLoadOptions options = OfficeTomlDocuments.BuildLoadOptions(Options, Untrusted.IsPresent);
        try {
            if (ParameterSetName == "Text") {
                WriteObject(TomlDocument.Parse(Text, options));
                return;
            }
            if (ParameterSetName == "Stream") {
                WriteObject(TomlDocument.Load(Stream!, options));
                return;
            }
        } catch (TomlException error) {
            ThrowTerminatingError(OfficeTomlDocuments.Error(error, ParameterSetName == "Text" ? Text : Stream));
        }
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string source in OfficePath.ResolveInput(SessionState, literal ? LiteralPath : Path, literal)) {
            try {
                RunDocument(source, isNew: false, force: false,
                    path => TomlDocument.Load(path!, options),
                    (document, path, _) => { document.Save(path); return null; });
            } catch (TomlException error) {
                WriteError(OfficeTomlDocuments.Error(error, source));
            }
        }
    }
}
