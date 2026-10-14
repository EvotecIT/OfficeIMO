using System.Collections.Generic;

namespace OfficeIMO.Chm;

/// <summary>A contents or keyword-index item, including book headings and multi-target keywords.</summary>
public sealed class ChmNavigationItem {
    internal ChmNavigationItem(string name, IReadOnlyList<ChmLink> links, IReadOnlyList<string> seeAlso, IReadOnlyList<ChmNavigationItem> children) {
        Name = name; Links = links; SeeAlso = seeAlso; Children = children;
    }
    /// <summary>Decoded display name.</summary>
    public string Name { get; }
    /// <summary>Zero or more topic targets. A book heading can have no targets.</summary>
    public IReadOnlyList<ChmLink> Links { get; }
    /// <summary>Keyword names referenced by See Also entries.</summary>
    public IReadOnlyList<string> SeeAlso { get; }
    /// <summary>Nested contents or index items, in source order.</summary>
    public IReadOnlyList<ChmNavigationItem> Children { get; }
}

/// <summary>A help-book navigation target. External and merged-help references are retained without loading them.</summary>
public sealed class ChmLink {
    internal ChmLink(string target, string? title = null) { Target = target; Title = title; }
    /// <summary>Local archive reference, including its fragment, or the original external reference.</summary>
    public string Target { get; }
    /// <summary>Optional per-target title, useful for multi-target index keywords.</summary>
    public string? Title { get; }
    /// <summary>Whether the reference requires a different archive or external resource.</summary>
    public bool IsExternal => Target.IndexOf("::", StringComparison.Ordinal) >= 0 ||
        (Uri.TryCreate(Target, UriKind.Absolute, out Uri? uri) && !string.IsNullOrEmpty(uri.Host)) ||
        (Target.IndexOf(':') >= 0 && !Target.StartsWith("/", StringComparison.Ordinal));
}
