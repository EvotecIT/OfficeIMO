using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets hyperlinks, filtered by text, URL or bookmark anchor.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeWordHyperlink", DefaultParameterSetName = "Document")]
[OutputType(typeof(WordHyperLink))]
public sealed class GetOfficeWordHyperlinkCommand : PSCmdlet {
    /// <summary>Path for this inspection.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Text for this inspection.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Section")]
    [Parameter(ParameterSetName = "Paragraph")]
    public string[]? Text { get; set; }
    /// <summary>Url for this inspection.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Section")]
    [Parameter(ParameterSetName = "Paragraph")]
    public string[]? Url { get; set; }
    /// <summary>Anchor for this inspection.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Section")]
    [Parameter(ParameterSetName = "Paragraph")]
    public string[]? Anchor { get; set; }
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
        if (ParameterSetName == "Section") { Emit(Section!.HyperLinks); return; }
        if (ParameterSetName == "Paragraph") { if (Paragraph!.Hyperlink != null) Emit(new[] { Paragraph.Hyperlink }); return; }
        WordQuerySource.Visit(this, ParameterSetName, Path, LiteralPath, Document, Inspect);
    }

    private void Inspect(WordDocument document) => Emit(document.HyperLinks);
    private void Emit(IEnumerable<WordHyperLink> items) {
        var text = WordFilterHelpers.BuildPatterns(Text);
        var url = WordFilterHelpers.BuildPatterns(Url);
        var anchor = WordFilterHelpers.BuildPatterns(Anchor);
        WriteObject(items.Where(item => WordFilterHelpers.Matches(item.Text, text) && WordFilterHelpers.Matches(item.Uri?.ToString(), url) && WordFilterHelpers.Matches(item.Anchor, anchor)), true);
    }
}
