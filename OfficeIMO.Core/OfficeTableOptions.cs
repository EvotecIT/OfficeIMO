using System;
using System.Collections.Generic;

namespace OfficeIMO;

/// <summary>Shared options for projecting typed rows into an editable native table.</summary>
public class OfficeTableOptions {
    /// <summary>Includes a first row containing the selected member names. Defaults to true.</summary>
    public bool IncludeHeaders { get; set; } = true;

    /// <summary>Optional native table name or accessible title.</summary>
    public string? Name { get; set; }

    /// <summary>Optional member names to include, matched without regard to case.</summary>
    public HashSet<string> Include { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Member names to exclude, matched without regard to case.</summary>
    public HashSet<string> Exclude { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Member names to place first. Remaining columns keep their discovery order.</summary>
    public List<string> Order { get; } = new();

    /// <summary>Display headers keyed by the original member name.</summary>
    public Dictionary<string, string> HeaderRenames { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Column alignments in output order. A null entry keeps the format's default alignment.</summary>
    public List<OfficeAlignment?> Alignments { get; } = new();

    /// <summary>Cell formatters keyed by the original member name. Formatter failures propagate to the caller.</summary>
    public Dictionary<string, Func<object?, string>> Formatters { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Transforms headers that have no explicit entry in <see cref="HeaderRenames"/>.</summary>
    public Func<string, string>? HeaderTransform { get; set; }
}
