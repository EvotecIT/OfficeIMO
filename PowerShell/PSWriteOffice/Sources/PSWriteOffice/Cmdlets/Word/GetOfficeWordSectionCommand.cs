using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets sections of a document.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeWordSection", DefaultParameterSetName = "Document")]
[OutputType(typeof(WordSection))]
public sealed class GetOfficeWordSectionCommand : PSCmdlet {
    /// <summary>Path for this inspection.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Index for this inspection.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public int[]? Index { get; set; }
    /// <summary>LiteralPath for this inspection.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Document for this inspection.</summary>
    [Parameter(ParameterSetName = "Document", ValueFromPipeline = true)]
    public WordDocument? Document { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordQuerySource.Visit(this, ParameterSetName, Path, LiteralPath, Document, Inspect);
    }

    private void Inspect(WordDocument document) {
        IEnumerable<WordSection> sections = document.Sections;
        if (Index != null && Index.Length > 0) {
            var selected = new List<WordSection>();
            foreach (int index in Index) {
                if (index < 1 || index > document.Sections.Count) throw new PSArgumentOutOfRangeException(nameof(Index), index, "Section indices start at one.");
                selected.Add(document.Sections[index - 1]);
            }
            sections = selected;
        }
        WriteObject(sections, true);
    }
}
