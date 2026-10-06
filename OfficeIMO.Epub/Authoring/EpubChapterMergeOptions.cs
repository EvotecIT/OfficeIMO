namespace OfficeIMO.Epub;

/// <summary>Controls stylesheet reconciliation when merging consecutive reflowable chapters.</summary>
public enum EpubChapterMergeStylePolicy {
    /// <summary>Require equivalent chapter heads after URL rebasing, apart from title text.</summary>
    RequireEquivalent,
    /// <summary>Retain first-head styles, then append every second-head style and stylesheet link in source order.
    /// Both chapters share the resulting cascade; no CSS isolation or visual equivalence is implied.</summary>
    AppendSecondStyles
}

/// <summary>Explicit reconciliation choices for chapter merges. Other structural conflicts remain errors.</summary>
public sealed class EpubChapterMergeOptions {
    /// <summary>
    /// Explicit replacements for second-chapter body identifiers, excluding shared merge containers.
    /// Fragment URLs and document-local relationships are repaired. CSS selectors are not rewritten;
    /// prepare ID-dependent styles for the selected names before merging. Fragment-only URLs in retained
    /// stylesheets naming changed IDs are rejected as ambiguous. At most 10,000 replacements.
    /// </summary>
    public IReadOnlyDictionary<string, string> SecondChapterIdMap { get; set; } = new Dictionary<string, string>();
    /// <summary>Style reconciliation policy. Defaults to rejecting different chapter heads.</summary>
    public EpubChapterMergeStylePolicy StylePolicy { get; set; }
}
