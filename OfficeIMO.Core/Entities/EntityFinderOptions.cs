using System;
using System.Collections.Generic;

namespace OfficeIMO.Entities;

/// <summary>Filters, scoring and input limits. A finder copies these settings at construction.</summary>
public sealed class EntityFinderOptions {
    /// <summary>Categories to include. An empty set includes all categories.</summary>
    public ISet<EntityKind> Kinds { get; } = new HashSet<EntityKind>();
    /// <summary>Countries to include. International schemes always qualify for this filter.</summary>
    public ISet<string> Countries { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    /// <summary>Scheme IDs to include. An empty set includes all IDs.</summary>
    public ISet<string> IncludeSchemes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    /// <summary>Scheme IDs to exclude.</summary>
    public ISet<string> ExcludeSchemes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    /// <summary>Keyword languages to use. Null enables every registered language.</summary>
    public IReadOnlyCollection<string>? Languages { get; set; }
    /// <summary>Lowest confidence included in the result.</summary>
    public double MinimumConfidence { get; set; } = 0.50;
    /// <summary>Maximum returned findings across segments.</summary>
    public int MaxMatches { get; set; } = 10_000;
    /// <summary>Maximum UTF-16 characters searched across segments.</summary>
    public int MaxInputCharacters { get; set; } = 16_777_216;
    /// <summary>Maximum distance before a value where a keyword can support it.</summary>
    public int ContextWindow { get; set; } = 48;
    /// <summary>Controls the returned readings when spans intersect.</summary>
    public EntityOverlapPolicy OverlapPolicy { get; set; } = EntityOverlapPolicy.BestOnly;
}
