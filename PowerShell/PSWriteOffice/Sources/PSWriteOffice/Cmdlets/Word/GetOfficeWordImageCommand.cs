using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets images from a document, section or paragraph.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeWordImage", DefaultParameterSetName = "Document")]
[OutputType(typeof(WordImage))]
public sealed class GetOfficeWordImageCommand : PSCmdlet {
    /// <summary>Path for this inspection.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>LiteralPath for this inspection.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Document for this inspection.</summary>
    [Parameter(ParameterSetName = "Document", ValueFromPipeline = true)]
    public WordDocument? Document { get; set; }
    /// <summary>Section for this inspection.</summary>
    [Parameter(ParameterSetName = "Section", Mandatory = true, ValueFromPipeline = true)]
    public WordSection? Section { get; set; }
    /// <summary>Paragraph for this inspection.</summary>
    [Parameter(ParameterSetName = "Paragraph", Mandatory = true, ValueFromPipeline = true)]
    public WordParagraph? Paragraph { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == "Section") { Emit(Section!.Images); return; }
        if (ParameterSetName == "Paragraph") { if (Paragraph!.Image != null) Emit(new[] { Paragraph.Image }); return; }
        WordQuerySource.Visit(this, ParameterSetName, Path, LiteralPath, Document, Inspect);
    }

    private void Inspect(WordDocument document) => Emit(document.Images);
    private void Emit(IEnumerable<WordImage> items) => WriteObject(items, true);
}
