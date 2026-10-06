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
    /// values are preserved. Attribute comparisons on document-local relationships and resource-bearing attributes follow actual
    /// repaired values, including rebased URLs and private stylesheet paths. Partial or multivalued matches can
    /// expand into at most 256 exact alternatives and 64 KiB of CSS, using :is when needed. Irreducible
    /// ambiguities and unsupported syntax fail atomically. Reader support for :is and nested CSS remains independent.</summary>
    public bool RewriteChapterSelectors { get; set; }
    /// <summary>Allow differing root/body language and direction by wrapping the second chapter's body content
    /// in a div with its effective language and direction. Other scaffold conflicts still reject the merge.
    /// Explicit auto direction on either source root/body is unsupported. Wrapper structure can affect CSS;
    /// this option does not isolate styles or establish equivalent reader layout.</summary>
    public bool PreserveSecondChapterLanguageAndDirection { get; set; }
    /// <summary>Style reconciliation policy. Defaults to rejecting different chapter heads.</summary>
    public EpubChapterMergeStylePolicy StylePolicy { get; set; }
}
