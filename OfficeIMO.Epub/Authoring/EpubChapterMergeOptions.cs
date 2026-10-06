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
    /// <summary>Style reconciliation policy. Defaults to rejecting different chapter heads.</summary>
    public EpubChapterMergeStylePolicy StylePolicy { get; set; }
}
