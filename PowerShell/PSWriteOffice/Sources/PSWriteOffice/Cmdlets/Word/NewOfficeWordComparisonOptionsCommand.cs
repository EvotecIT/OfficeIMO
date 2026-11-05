using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates discoverable structural comparison settings for Compare-OfficeDocument.</summary>
/// <example>
///   <summary>Ignore text normalization differences and exclude volatile metadata.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeWordComparisonOptions -IgnoreWhitespace -IgnoreCase -NoVolatileMetadata:$false
/// Compare-OfficeDocument -ReferencePath .\Before.docx -DifferencePath .\After.docx -Options $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeWordComparisonOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordComparisonOptions))]
public sealed class NewOfficeWordComparisonOptionsCommand : PSCmdlet {
    /// <summary>Ignore differences caused only by whitespace runs.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter IgnoreWhitespace { get; set; }
    /// <summary>Ignore character casing.</summary>
    [Parameter] public SwitchParameter IgnoreCase { get; set; }
    /// <summary>Compare direct run formatting.</summary>
    [Parameter] public SwitchParameter NoRunFormatting { get; set; }
    /// <summary>Compare resolved effective formatting.</summary>
    [Parameter] public SwitchParameter NoEffectiveFormatting { get; set; }
    /// <summary>Compare paragraph style identifiers.</summary>
    [Parameter] public SwitchParameter NoParagraphStyleIds { get; set; }
    /// <summary>Compare run style identifiers.</summary>
    [Parameter] public SwitchParameter NoRunStyleIds { get; set; }
    /// <summary>Limit results to these comparison scopes.</summary>
    [Parameter] public WordComparisonScope[]? IncludeScope { get; set; }
    /// <summary>Remove these comparison scopes from results.</summary>
    [Parameter] public WordComparisonScope[]? ExcludeScope { get; set; }
    /// <summary>Compare fields.</summary>
    [Parameter] public SwitchParameter NoFields { get; set; }
    /// <summary>Compare content controls.</summary>
    [Parameter] public SwitchParameter NoContentControls { get; set; }
    /// <summary>Compare bookmarks.</summary>
    [Parameter] public SwitchParameter NoBookmarks { get; set; }
    /// <summary>Compare hyperlinks.</summary>
    [Parameter] public SwitchParameter NoHyperlinks { get; set; }
    /// <summary>Compare lists.</summary>
    [Parameter] public SwitchParameter NoLists { get; set; }
    /// <summary>Compare comments.</summary>
    [Parameter] public SwitchParameter NoComments { get; set; }
    /// <summary>Compare comment authors.</summary>
    [Parameter] public SwitchParameter NoCommentAuthors { get; set; }
    /// <summary>Compare comment text.</summary>
    [Parameter] public SwitchParameter NoCommentText { get; set; }
    /// <summary>Compare comment resolved state.</summary>
    [Parameter] public SwitchParameter NoCommentResolvedState { get; set; }
    /// <summary>Compare comment targets.</summary>
    [Parameter] public SwitchParameter NoCommentTargets { get; set; }
    /// <summary>Compare comment replies.</summary>
    [Parameter] public SwitchParameter NoCommentReplies { get; set; }
    /// <summary>Compare tracked revisions.</summary>
    [Parameter] public SwitchParameter NoRevisions { get; set; }
    /// <summary>Compare revision authors.</summary>
    [Parameter] public SwitchParameter NoRevisionAuthors { get; set; }
    /// <summary>Compare revision text.</summary>
    [Parameter] public SwitchParameter NoRevisionText { get; set; }
    /// <summary>Compare revision locations.</summary>
    [Parameter] public SwitchParameter NoRevisionLocations { get; set; }
    /// <summary>Compare images.</summary>
    [Parameter] public SwitchParameter NoImages { get; set; }
    /// <summary>Compare supported shapes.</summary>
    [Parameter] public SwitchParameter NoShapes { get; set; }
    /// <summary>Compare document block order.</summary>
    [Parameter] public SwitchParameter NoBlockOrder { get; set; }
    /// <summary>Compare generated identifiers.</summary>
    [Parameter] public SwitchParameter NoGeneratedIds { get; set; }
    /// <summary>Compare volatile timestamps and metadata.</summary>
    [Parameter] public SwitchParameter NoVolatileMetadata { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new WordComparisonOptions();
        Apply(nameof(IgnoreWhitespace), value => options.IgnoreWhitespace = value);
        Apply(nameof(IgnoreCase), value => options.IgnoreCase = value);
        Apply(nameof(NoRunFormatting), value => options.CompareRunFormatting = !value);
        Apply(nameof(NoEffectiveFormatting), value => options.CompareEffectiveFormatting = !value);
        Apply(nameof(NoParagraphStyleIds), value => options.CompareParagraphStyleIds = !value);
        Apply(nameof(NoRunStyleIds), value => options.CompareRunStyleIds = !value);
        Apply(nameof(NoFields), value => options.CompareFields = !value);
        Apply(nameof(NoContentControls), value => options.CompareContentControls = !value);
        Apply(nameof(NoBookmarks), value => options.CompareBookmarks = !value);
        Apply(nameof(NoHyperlinks), value => options.CompareHyperlinks = !value);
        Apply(nameof(NoLists), value => options.CompareLists = !value);
        Apply(nameof(NoComments), value => options.CompareComments = !value);
        Apply(nameof(NoCommentAuthors), value => options.CompareCommentAuthors = !value);
        Apply(nameof(NoCommentText), value => options.CompareCommentText = !value);
        Apply(nameof(NoCommentResolvedState), value => options.CompareCommentResolvedState = !value);
        Apply(nameof(NoCommentTargets), value => options.CompareCommentTargets = !value);
        Apply(nameof(NoCommentReplies), value => options.CompareCommentReplies = !value);
        Apply(nameof(NoRevisions), value => options.CompareRevisions = !value);
        Apply(nameof(NoRevisionAuthors), value => options.CompareRevisionAuthors = !value);
        Apply(nameof(NoRevisionText), value => options.CompareRevisionText = !value);
        Apply(nameof(NoRevisionLocations), value => options.CompareRevisionLocations = !value);
        Apply(nameof(NoImages), value => options.CompareImages = !value);
        Apply(nameof(NoShapes), value => options.CompareShapes = !value);
        Apply(nameof(NoBlockOrder), value => options.CompareBlockOrder = !value);
        Apply(nameof(NoGeneratedIds), value => options.CompareGeneratedIds = !value);
        Apply(nameof(NoVolatileMetadata), value => options.CompareVolatileMetadata = !value);
        if (IncludeScope != null) options.IncludedScopes = new HashSet<WordComparisonScope>(IncludeScope);
        if (ExcludeScope != null) options.ExcludedScopes = new HashSet<WordComparisonScope>(ExcludeScope);
        WriteObject(options);
    }

    private void Apply(string name, System.Action<bool> setter) {
        if (!MyInvocation.BoundParameters.ContainsKey(name)) return;
        setter(((SwitchParameter)MyInvocation.BoundParameters[name]).IsPresent);
    }
}
