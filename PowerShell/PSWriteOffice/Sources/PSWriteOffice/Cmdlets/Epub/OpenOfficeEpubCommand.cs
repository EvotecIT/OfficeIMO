using System.Management.Automation;
using OfficeIMO.Epub;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Epub;

/// <summary>Opens an EPUB publication, optionally saving an editing block atomically.</summary>
/// <example><summary>Open a publication from a literal file pipeline.</summary>
/// <code>$book = Get-Item -LiteralPath './Guide[final].epub' | Open-OfficeEpub</code></example>
/// <example><summary>Append a chapter and save only when the block succeeds.</summary>
/// <code>Open-OfficeEpub -Path ./Guide.epub { EpubChapter -Title Review -Text 'Reviewed and approved.' }</code></example>
[Cmdlet(VerbsCommon.Open, "OfficeEpub", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(EpubPublication), typeof(FileInfo))]
public sealed class OpenOfficeEpubCommand : OfficeLifecycleCmdlet<EpubPublication> {
    /// <summary>An optional authoring block; omission returns the live document.</summary>
    [Parameter(ParameterSetName = "Path", Position = 1)]
    [Parameter(ParameterSetName = "LiteralPath", Position = 1)]
    [ValidateNotNull]
    public override ScriptBlock? Content { get => base.Content; set => base.Content = value; }

    /// <summary>Emit the FileInfo after a completed authoring block.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public override SwitchParameter PassThru { get => base.PassThru; set => base.PassThru = value; }

    /// <summary>Explicitly accept native-save loss for an authoring block.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public override SwitchParameter AllowLoss { get => base.AllowLoss; set => base.AllowLoss = value; }

    /// <summary>Open the saved document after a completed authoring block.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public override SwitchParameter Show { get => base.Show; set => base.Show = value; }

    /// <summary>EPUB source with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;

    /// <summary>Literal EPUB source, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string LiteralPath { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string source in OfficePath.ResolveInput(SessionState, literal ? LiteralPath : Path, literal)) {
            RunDocument(source, isNew: false, force: false, path => EpubPublication.Load(path!),
                (document, path, allowLoss) => {
                    EpubWriteReport report = document.Save(path);
                    if (!allowLoss) report.ThrowIfLoss();
                    return null;
                });
        }
    }
}
