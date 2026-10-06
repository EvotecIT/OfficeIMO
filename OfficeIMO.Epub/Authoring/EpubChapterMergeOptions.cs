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
    /// Fragment URLs and document-local relationships are repaired. CSS selectors remain unchanged unless
    /// RewriteChapterSelectors is enabled. Fragment-only URLs in retained
    /// stylesheets naming changed IDs are rejected as ambiguous. At most 10,000 replacements.
    /// </summary>
    public IReadOnlyDictionary<string, string> SecondChapterIdMap { get; set; } = new Dictionary<string, string>();
    /// <summary>Repair selectors in both source chapters, applying the identifier map only to the second chapter
    /// and retaining separate private copies of each chapter's linked stylesheets and imports.
    /// Requires AppendSecondStyles. Supports hash and exact id attribute selectors
    /// in ordinary/nested rules and media/supports/layer/container/scope groups. Declaration order and custom-property
    /// values are preserved. Exact and whitespace-token selectors on document-local relationships follow actual
    /// attribute changes. Exact and presence selectors on resource-bearing attributes follow rebased URLs and final
    /// stylesheet clone paths, including merges without identifier replacements. Ambiguous replacements, partial
    /// URL matches and unsupported syntax fail atomically. Reader support for nested CSS remains independent.</summary>
    public bool RewriteChapterSelectors { get; set; }
    /// <summary>Style reconciliation policy. Defaults to rejecting different chapter heads.</summary>
    public EpubChapterMergeStylePolicy StylePolicy { get; set; }
}
