using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets bookmarks, optionally by name (wildcards).</summary>
[Cmdlet(VerbsCommon.Get, "OfficeWordBookmark", DefaultParameterSetName = "Document")]
[OutputType(typeof(WordBookmark))]
public sealed class GetOfficeWordBookmarkCommand : PSCmdlet {
    /// <summary>Path for this inspection.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Name for this inspection.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public string[]? Name { get; set; }
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
        var names = WordFilterHelpers.BuildPatterns(Name);
        WriteObject(document.Bookmarks.Where(item => WordFilterHelpers.Matches(item.Name, names)), true);
    }
}
