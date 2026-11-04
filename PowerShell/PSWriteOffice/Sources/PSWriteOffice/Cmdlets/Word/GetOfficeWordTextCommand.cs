using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets the text runs of paragraphs, sections or documents.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeWordText", DefaultParameterSetName = "Document")]
[OutputType(typeof(WordParagraph))]
public sealed class GetOfficeWordTextCommand : PSCmdlet {
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
        if (ParameterSetName == "Section") { WriteParagraphs(Section!.Paragraphs); return; }
        if (ParameterSetName == "Paragraph") { WriteParagraphs(new[] { Paragraph! }); return; }
        WordQuerySource.Visit(this, ParameterSetName, Path, LiteralPath, Document, Inspect);
    }

    private void Inspect(WordDocument document) => WriteParagraphs(document.Paragraphs);
    private void WriteParagraphs(IEnumerable<WordParagraph> paragraphs) {
        foreach (var paragraph in paragraphs) WriteObject(paragraph.GetRuns(), true);
    }
}
