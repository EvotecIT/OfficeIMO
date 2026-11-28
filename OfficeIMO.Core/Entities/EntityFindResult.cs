using System;
using System.Collections.Generic;

namespace OfficeIMO.Entities;

/// <summary>The findings and value-free diagnostics for a find operation.</summary>
public sealed class EntityFindResult {
    internal EntityFindResult(List<EntityMatch> matches, bool truncated, int droppedOverlaps) {
        Matches = matches.AsReadOnly();
        Truncated = truncated;
        DroppedOverlaps = droppedOverlaps;
        var diagnostics = new List<string>(2);
        if (truncated) diagnostics.Add("The configured input or match limit was reached.");
        if (droppedOverlaps > 0) diagnostics.Add("Overlapping entity readings were omitted.");
        Diagnostics = diagnostics.AsReadOnly();
    }

    /// <summary>Findings in segment and source order.</summary>
    public IReadOnlyList<EntityMatch> Matches { get; }
    /// <summary>Whether a limit prevented returning all findings or searching all characters.</summary>
    public bool Truncated { get; }
    /// <summary>Number of otherwise eligible readings dropped because they overlapped a winner.</summary>
    public int DroppedOverlaps { get; }
    /// <summary>Privacy-preserving descriptions of the limits encountered.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
}
