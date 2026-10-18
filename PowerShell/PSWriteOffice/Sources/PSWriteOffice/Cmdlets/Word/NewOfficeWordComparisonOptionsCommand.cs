using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates discoverable structural comparison settings for Compare-OfficeWordDocument.</summary>
/// <example>
///   <summary>Ignore text normalization differences and exclude volatile metadata.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeWordComparisonOptions -IgnoreWhitespace -IgnoreCase -CompareVolatileMetadata:$false
/// Compare-OfficeWordDocument -ReferencePath .\Before.docx -DifferencePath .\After.docx -Options $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeWordComparisonOptions")]
[OutputType(typeof(WordComparisonOptions))]
public sealed class NewOfficeWordComparisonOptionsCommand : PSCmdlet {
    /// <summary>Ignore differences caused only by whitespace runs.</summary>
    [Parameter] public SwitchParameter IgnoreWhitespace { get; set; }
    /// <summary>Ignore character casing.</summary>
    [Parameter] public SwitchParameter IgnoreCase { get; set; }
    /// <summary>Compare direct run formatting.</summary>
    [Parameter] public SwitchParameter CompareRunFormatting { get; set; }
    /// <summary>Compare resolved effective formatting.</summary>
    [Parameter] public SwitchParameter CompareEffectiveFormatting { get; set; }
    /// <summary>Compare paragraph style identifiers.</summary>
    [Parameter] public SwitchParameter CompareParagraphStyleIds { get; set; }
    /// <summary>Compare run style identifiers.</summary>
    [Parameter] public SwitchParameter CompareRunStyleIds { get; set; }
    /// <summary>Limit results to these comparison scopes.</summary>
    [Parameter] public WordComparisonScope[]? IncludeScope { get; set; }
    /// <summary>Remove these comparison scopes from results.</summary>
    [Parameter] public WordComparisonScope[]? ExcludeScope { get; set; }
    /// <summary>Compare fields.</summary>
    [Parameter] public SwitchParameter CompareFields { get; set; }
    /// <summary>Compare content controls.</summary>
    [Parameter] public SwitchParameter CompareContentControls { get; set; }
    /// <summary>Compare bookmarks.</summary>
    [Parameter] public SwitchParameter CompareBookmarks { get; set; }
    /// <summary>Compare hyperlinks.</summary>
    [Parameter] public SwitchParameter CompareHyperlinks { get; set; }
    /// <summary>Compare lists.</summary>
    [Parameter] public SwitchParameter CompareLists { get; set; }
    /// <summary>Compare comments.</summary>
    [Parameter] public SwitchParameter CompareComments { get; set; }
    /// <summary>Compare comment authors.</summary>
    [Parameter] public SwitchParameter CompareCommentAuthors { get; set; }
    /// <summary>Compare comment text.</summary>
    [Parameter] public SwitchParameter CompareCommentText { get; set; }
    /// <summary>Compare comment resolved state.</summary>
    [Parameter] public SwitchParameter CompareCommentResolvedState { get; set; }
    /// <summary>Compare comment targets.</summary>
    [Parameter] public SwitchParameter CompareCommentTargets { get; set; }
    /// <summary>Compare comment replies.</summary>
    [Parameter] public SwitchParameter CompareCommentReplies { get; set; }
    /// <summary>Compare tracked revisions.</summary>
    [Parameter] public SwitchParameter CompareRevisions { get; set; }
    /// <summary>Compare revision authors.</summary>
    [Parameter] public SwitchParameter CompareRevisionAuthors { get; set; }
    /// <summary>Compare revision text.</summary>
    [Parameter] public SwitchParameter CompareRevisionText { get; set; }
    /// <summary>Compare revision locations.</summary>
    [Parameter] public SwitchParameter CompareRevisionLocations { get; set; }
    /// <summary>Compare images.</summary>
    [Parameter] public SwitchParameter CompareImages { get; set; }
    /// <summary>Compare supported shapes.</summary>
    [Parameter] public SwitchParameter CompareShapes { get; set; }
    /// <summary>Compare document block order.</summary>
    [Parameter] public SwitchParameter CompareBlockOrder { get; set; }
    /// <summary>Compare generated identifiers.</summary>
    [Parameter] public SwitchParameter CompareGeneratedIds { get; set; }
    /// <summary>Compare volatile timestamps and metadata.</summary>
    [Parameter] public SwitchParameter CompareVolatileMetadata { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new WordComparisonOptions();
        Apply(nameof(IgnoreWhitespace), value => options.IgnoreWhitespace = value);
        Apply(nameof(IgnoreCase), value => options.IgnoreCase = value);
        Apply(nameof(CompareRunFormatting), value => options.CompareRunFormatting = value);
        Apply(nameof(CompareEffectiveFormatting), value => options.CompareEffectiveFormatting = value);
        Apply(nameof(CompareParagraphStyleIds), value => options.CompareParagraphStyleIds = value);
        Apply(nameof(CompareRunStyleIds), value => options.CompareRunStyleIds = value);
        Apply(nameof(CompareFields), value => options.CompareFields = value);
        Apply(nameof(CompareContentControls), value => options.CompareContentControls = value);
        Apply(nameof(CompareBookmarks), value => options.CompareBookmarks = value);
        Apply(nameof(CompareHyperlinks), value => options.CompareHyperlinks = value);
        Apply(nameof(CompareLists), value => options.CompareLists = value);
        Apply(nameof(CompareComments), value => options.CompareComments = value);
        Apply(nameof(CompareCommentAuthors), value => options.CompareCommentAuthors = value);
        Apply(nameof(CompareCommentText), value => options.CompareCommentText = value);
        Apply(nameof(CompareCommentResolvedState), value => options.CompareCommentResolvedState = value);
        Apply(nameof(CompareCommentTargets), value => options.CompareCommentTargets = value);
        Apply(nameof(CompareCommentReplies), value => options.CompareCommentReplies = value);
        Apply(nameof(CompareRevisions), value => options.CompareRevisions = value);
        Apply(nameof(CompareRevisionAuthors), value => options.CompareRevisionAuthors = value);
        Apply(nameof(CompareRevisionText), value => options.CompareRevisionText = value);
        Apply(nameof(CompareRevisionLocations), value => options.CompareRevisionLocations = value);
        Apply(nameof(CompareImages), value => options.CompareImages = value);
        Apply(nameof(CompareShapes), value => options.CompareShapes = value);
        Apply(nameof(CompareBlockOrder), value => options.CompareBlockOrder = value);
        Apply(nameof(CompareGeneratedIds), value => options.CompareGeneratedIds = value);
        Apply(nameof(CompareVolatileMetadata), value => options.CompareVolatileMetadata = value);
        if (IncludeScope != null) options.IncludedScopes = new HashSet<WordComparisonScope>(IncludeScope);
        if (ExcludeScope != null) options.ExcludedScopes = new HashSet<WordComparisonScope>(ExcludeScope);
        WriteObject(options);
    }

    private void Apply(string name, System.Action<bool> setter) {
        if (!MyInvocation.BoundParameters.ContainsKey(name)) return;
        setter(((SwitchParameter)MyInvocation.BoundParameters[name]).IsPresent);
    }
}
