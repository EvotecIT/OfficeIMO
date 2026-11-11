using System.Management.Automation;
using OfficeIMO.Yaml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Yaml;

namespace PSWriteOffice.Cmdlets.Yaml;

/// <summary>Opens a YAML file, stream or text as a lossless live document; a file content block saves atomically.</summary>
/// <para>Saving writes only what changed: comments, blank lines, quoting and indentation outside the edited nodes stay
/// byte for byte. Malformed input fails with a diagnostic error ID (for example yaml-syntax-unexpected-token), and the
/// message gives the line, column and offending source line.</para>
/// <example><summary>Change one value and keep the rest of the file as it was.</summary>
/// <code>Open-OfficeYaml -Path ./deployment.yaml {
///     Set-OfficeYamlNode -Select 'spec.replicas' -Value 3
/// }</code></example>
/// <example><summary>Read an uploaded file with untrusted limits.</summary>
/// <code>$document = Open-OfficeYaml -LiteralPath ./uploads/values.yaml -Untrusted -Strict
/// Get-OfficeYamlNode -Document $document -Select 'image.tag'
/// Close-OfficeYaml -Document $document</code></example>
[Cmdlet(VerbsCommon.Open, "OfficeYaml", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(YamlDocument), typeof(FileInfo))]
public sealed class OpenOfficeYamlCommand : OfficeLifecycleCmdlet<YamlDocument> {
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

    /// <summary>YAML file with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;
    /// <summary>Literal YAML file, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string LiteralPath { get; set; } = string.Empty;
    /// <summary>YAML text to parse without a file.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>A readable stream, read from its current position; it stays open.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Stream"), ValidateNotNull]
    public Stream? Stream { get; set; }
    /// <summary>Native load options or a hashtable of their properties: limits, schema, strictness and duplicate keys.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(YamlLoadOptions)), ValidateNotNull]
    public YamlLoadOptions? Options { get; set; }
    /// <summary>Caps every limit at the engine's untrusted defaults, for files from outside.</summary>
    [Parameter]
    public SwitchParameter Untrusted { get; set; }
    /// <summary>Applies the YAML specification's exact rules instead of accepting common real-world leniencies.</summary>
    [Parameter]
    public SwitchParameter Strict { get; set; }
    /// <summary>Forces a schema for typed values: Failsafe, Json, Core (the default) or Yaml11.</summary>
    [Parameter]
    public YamlSchema? Schema { get; set; }
    /// <summary>Native save settings used when an editing block saves; omission writes a minimal-change save.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "LiteralPath")]
    [OfficeOptionsTransformation(typeof(YamlSaveOptions)), ValidateNotNull]
    public YamlSaveOptions? WriteOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        YamlLoadOptions options = OfficeYamlDocuments.BuildLoadOptions(Options, Untrusted.IsPresent, BoundSwitch(nameof(Strict)), Schema);
        try {
            if (ParameterSetName == "Text") {
                WriteObject(YamlDocument.Parse(Text, options));
                return;
            }
            if (ParameterSetName == "Stream") {
                WriteObject(YamlDocument.Load(Stream!, options));
                return;
            }
        } catch (YamlException error) {
            ThrowTerminatingError(OfficeYamlDocuments.Error(error, ParameterSetName == "Text" ? Text : Stream));
        }
        if (Content == null && WriteOptions != null) {
            InvalidArgument(new PSArgumentException("WriteOptions require a content block; use Save-OfficeYaml for a live document."), Path);
            return;
        }
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string source in OfficePath.ResolveInput(SessionState, literal ? LiteralPath : Path, literal)) {
            try {
                RunDocument(source, isNew: false, force: false,
                    path => OfficeYamlDocuments.Associate(YamlDocument.Load(path!, options), path),
                    (document, path, _) => { document.Save(path, WriteOptions); return null; });
            } catch (YamlException error) {
                WriteError(OfficeYamlDocuments.Error(error, source));
            }
        }
    }
}
