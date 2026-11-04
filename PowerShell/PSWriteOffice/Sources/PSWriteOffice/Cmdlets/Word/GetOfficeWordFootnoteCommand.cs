using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets footnotes of a document or section.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeWordFootnote", DefaultParameterSetName = "Document")]
[OutputType(typeof(WordNoteInfo))]
public sealed class GetOfficeWordFootnoteCommand : PSCmdlet {
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

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == "Section") { WriteObject(Section!.FootNotes.Select(ToInfo), true); return; }
        WordQuerySource.Visit(this, ParameterSetName, Path, LiteralPath, Document, Inspect);
    }

    private void Inspect(WordDocument document) => WriteObject(document.Sections.SelectMany(section => section.FootNotes).Select(ToInfo), true);
    private static WordNoteInfo ToInfo(WordFootNote note) {
        var paragraphs = note.Paragraphs?
            .Select(paragraph => paragraph.Text)
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToArray() ?? Array.Empty<string>();

        return new WordNoteInfo("Footnote", note.ReferenceId, note.ParentParagraph?.Text, paragraphs);
    }
}
