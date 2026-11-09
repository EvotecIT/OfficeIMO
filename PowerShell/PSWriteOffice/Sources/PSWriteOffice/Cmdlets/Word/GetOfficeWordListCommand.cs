using System.Management.Automation;
using System.Text.RegularExpressions;
using OfficeIMO;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets lists, optionally only those with an item matching a pattern.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeWordList", DefaultParameterSetName = "Document")]
[OutputType(typeof(WordList))]
public sealed class GetOfficeWordListCommand : PSCmdlet {
    /// <summary>Document paths; supports wildcards.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Include list definitions with no items.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Section")]
    public SwitchParameter IncludeEmpty { get; set; }

    /// <summary>Optional regular expression, or literal text with SimpleMatch.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Section")]
    public string? Pattern { get; set; }

    /// <summary>Treat Pattern as literal text.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Section")]
    public SwitchParameter SimpleMatch { get; set; }

    /// <summary>Use case-sensitive matching.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Section")]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Exact document paths, including literal brackets.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Caller-owned document; omission uses the enclosing Word content block.</summary>
    [Parameter(ParameterSetName = "Document", ValueFromPipeline = true)]
    public WordDocument? Document { get; set; }

    /// <summary>Caller-owned section to inspect.</summary>
    [Parameter(ParameterSetName = "Section", Mandatory = true, ValueFromPipeline = true)]
    public WordSection? Section { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == "Section") InspectSection(Section!);
        else WordQuerySource.Visit(this, ParameterSetName, Path, LiteralPath, Document, InspectDocument);
    }

    private void InspectDocument(WordDocument document) => Emit(document.Lists);
    private void InspectSection(WordSection section) => Emit(section.Lists);

    private void Emit(IEnumerable<WordList> items) {
        if (!IncludeEmpty) items = items.Where(item => item.ListItems.Count > 0);
        if (Pattern != null) {
            var matches = SimpleMatch ? WordObjectSearch.CreateTextMatcher(Pattern, CaseSensitive)
                : WordObjectSearch.CreateRegexMatcher(Pattern, CaseSensitive);
            items = items.Where(item => WordObjectSearch.MatchesList(item, matches));
        }
        WriteObject(items, true);
    }
}
