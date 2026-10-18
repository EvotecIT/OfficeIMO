using System.Management.Automation;
using OfficeIMO.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Creates discoverable token and hierarchy settings for Get-OfficeDocumentHierarchy.</summary>
/// <example>
///   <summary>Create embedding-ready chunks.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeReaderHierarchyOptions -MaxTokens 500 -OverlapTokens 50 -IncludeContextInText
/// Get-OfficeDocumentHierarchy -Path .\handbook.pdf -ChunkingOptions $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeReaderHierarchyOptions")]
[OutputType(typeof(ReaderHierarchicalChunkingOptions))]
public sealed class NewOfficeReaderHierarchyOptionsCommand : PSCmdlet {
    /// <summary>Maximum tokens per output chunk.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaxTokens { get; set; }
    /// <summary>Tokens repeated between adjacent chunks.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int? OverlapTokens { get; set; }
    /// <summary>Maximum source chunks accepted.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaxInputChunks { get; set; }
    /// <summary>Maximum chunks returned.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaxOutputChunks { get; set; }
    /// <summary>Maximum heading hierarchy depth.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaxHierarchyDepth { get; set; }
    /// <summary>Maximum heading-context characters retained.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int? MaxContextCharacters { get; set; }
    /// <summary>Prefer Markdown text where the reader supports it.</summary>
    [Parameter] public SwitchParameter PreferMarkdown { get; set; }
    /// <summary>Include hierarchy context in chunk text.</summary>
    [Parameter] public SwitchParameter IncludeContextInText { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new ReaderHierarchicalChunkingOptions();
        if (MaxTokens.HasValue) options.MaxTokens = MaxTokens.Value;
        if (OverlapTokens.HasValue) options.OverlapTokens = OverlapTokens.Value;
        if (MaxInputChunks.HasValue) options.MaxInputChunks = MaxInputChunks.Value;
        if (MaxOutputChunks.HasValue) options.MaxOutputChunks = MaxOutputChunks.Value;
        if (MaxHierarchyDepth.HasValue) options.MaxHierarchyDepth = MaxHierarchyDepth.Value;
        if (MaxContextCharacters.HasValue) options.MaxContextCharacters = MaxContextCharacters.Value;
        if (IsBound(nameof(PreferMarkdown))) options.PreferMarkdown = PreferMarkdown.IsPresent;
        if (IsBound(nameof(IncludeContextInText))) options.IncludeContextInText = IncludeContextInText.IsPresent;
        WriteObject(options);
    }

    private bool IsBound(string name) => MyInvocation.BoundParameters.ContainsKey(name);
}
