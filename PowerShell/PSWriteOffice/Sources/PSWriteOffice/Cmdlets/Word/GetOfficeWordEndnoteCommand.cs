using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets endnotes of a document or section.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeWordEndnote", DefaultParameterSetName = "Document")]
[OutputType(typeof(WordNoteInfo))]
public sealed class GetOfficeWordEndnoteCommand : PSCmdlet {
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
        if (ParameterSetName == "Section") { WriteObject(Section!.EndNotes.Select(ToInfo), true); return; }
        WordQuerySource.Visit(this, ParameterSetName, Path, LiteralPath, Document, Inspect);
    }

    private void Inspect(WordDocument document) => WriteObject(document.Sections.SelectMany(section => section.EndNotes).Select(ToInfo), true);
    private static WordNoteInfo ToInfo(WordEndNote note) {
        return new WordNoteInfo("Endnote", note.ReferenceId,
            string.Concat(note.Paragraph.GetRuns().Select(run => run.Text)),
            note.TextParagraphs ?? Array.Empty<string>());
    }
}
