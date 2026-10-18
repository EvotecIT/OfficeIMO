using System.Management.Automation;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Creates bounded token-aware chunks and a deterministic document hierarchy.</summary>
/// <example>
///   <summary>Create embedding-ready chunks with heading context.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeReaderHierarchyOptions -MaxTokens 500 -OverlapTokens 50 -IncludeContextInText
/// $result = Get-OfficeDocumentHierarchy -Path .\handbook.pdf -ChunkingOptions $options</code>
///   <para>Returns chunks, token evidence, overlap counts, and flattened parent/child nodes.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentHierarchy")]
[OutputType(typeof(ReaderChunkHierarchyResult))]
public sealed class GetOfficeDocumentHierarchyCommand : OfficeDocumentReaderCommandBase
{
    /// <summary>Path to read.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional source-reading limits and format behavior.</summary>
    [Parameter]
    public ReaderOptions? ReaderOptions { get; set; }

    /// <summary>Optional token budget, overlap, hierarchy, and token-counter settings.</summary>
    [Parameter]
    public ReaderHierarchicalChunkingOptions? ChunkingOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() => WriteObject(EffectiveReader.ReadHierarchical(
        ReaderCommandUtilities.ResolvePath(this, Path), ReaderOptions, ChunkingOptions));
}
