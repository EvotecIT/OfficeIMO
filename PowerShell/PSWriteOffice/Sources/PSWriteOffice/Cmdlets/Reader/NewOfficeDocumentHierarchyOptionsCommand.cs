using System.Management.Automation;
using OfficeIMO.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Creates discoverable token and hierarchy settings for Read-OfficeDocument -As Hierarchy.</summary>
/// <example>
///   <summary>Create embedding-ready chunks.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeDocumentHierarchyOptions -MaximumTokens 500 -OverlapTokens 50 -IncludeContextInText
/// Read-OfficeDocument -As Hierarchy -Path .\handbook.pdf -HierarchyOptions $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeDocumentHierarchyOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(ReaderHierarchicalChunkingOptions))]
public sealed class NewOfficeDocumentHierarchyOptionsCommand : PSCmdlet {
    /// <summary>Maximum tokens per output chunk.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaximumTokens { get; set; }
    /// <summary>Tokens repeated between adjacent chunks.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int? OverlapTokens { get; set; }
    /// <summary>Maximum source chunks accepted.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaximumInputChunks { get; set; }
    /// <summary>Maximum chunks returned.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaximumOutputChunks { get; set; }
    /// <summary>Maximum heading hierarchy depth.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaximumHierarchyDepth { get; set; }
    /// <summary>Maximum heading-context characters retained.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int? MaximumContextCharacters { get; set; }
    /// <summary>Prefer Markdown text where the reader supports it.</summary>
    [Parameter] public SwitchParameter PreferMarkdown { get; set; }
    /// <summary>Include hierarchy context in chunk text.</summary>
    [Parameter] public SwitchParameter IncludeContextInText { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new ReaderHierarchicalChunkingOptions();
        if (MaximumTokens.HasValue) options.MaxTokens = MaximumTokens.Value;
        if (OverlapTokens.HasValue) options.OverlapTokens = OverlapTokens.Value;
        if (MaximumInputChunks.HasValue) options.MaxInputChunks = MaximumInputChunks.Value;
        if (MaximumOutputChunks.HasValue) options.MaxOutputChunks = MaximumOutputChunks.Value;
        if (MaximumHierarchyDepth.HasValue) options.MaxHierarchyDepth = MaximumHierarchyDepth.Value;
        if (MaximumContextCharacters.HasValue) options.MaxContextCharacters = MaximumContextCharacters.Value;
        if (IsBound(nameof(PreferMarkdown))) options.PreferMarkdown = PreferMarkdown.IsPresent;
        if (IsBound(nameof(IncludeContextInText))) options.IncludeContextInText = IncludeContextInText.IsPresent;
        WriteObject(options);
    }

    private bool IsBound(string name) => MyInvocation.BoundParameters.ContainsKey(name);
}
