using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>Identifies a semantic diagram layout supported by the shared drawing renderer.</summary>
public enum OfficeDiagramKind {
    /// <summary>Sequential process or list nodes.</summary>
    Process,

    /// <summary>Parent-child hierarchy nodes.</summary>
    Hierarchy,

    /// <summary>Circular sequence nodes.</summary>
    Cycle
}

/// <summary>Dependency-free semantic diagram data for static rendering and export.</summary>
public sealed class OfficeDiagramSnapshot {
    /// <summary>Creates a semantic diagram snapshot.</summary>
    public OfficeDiagramSnapshot(string? name, OfficeDiagramKind kind,
        IEnumerable<string> nodes, double widthPoints,
        double heightPoints) {
        if (nodes == null) throw new ArgumentNullException(nameof(nodes));
        if (double.IsNaN(widthPoints) || double.IsInfinity(widthPoints)
            || widthPoints <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(widthPoints));
        }
        if (double.IsNaN(heightPoints) || double.IsInfinity(heightPoints)
            || heightPoints <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(heightPoints));
        }
        var values = new List<string>();
        foreach (string? node in nodes) {
            string value = (node ?? string.Empty).Trim();
            if (value.Length > 0) values.Add(value);
        }
        if (values.Count == 0) {
            throw new ArgumentException(
                "A diagram snapshot requires at least one non-empty node.",
                nameof(nodes));
        }
        if (values.Count > 4096) {
            throw new ArgumentException(
                "A diagram snapshot supports at most 4,096 nodes.",
                nameof(nodes));
        }
        Name = name;
        Kind = kind;
        Nodes = new ReadOnlyCollection<string>(values);
        WidthPoints = widthPoints;
        HeightPoints = heightPoints;
    }

    /// <summary>Gets the optional source diagram name.</summary>
    public string? Name { get; }

    /// <summary>Gets the semantic layout kind.</summary>
    public OfficeDiagramKind Kind { get; }

    /// <summary>Gets node labels in semantic order.</summary>
    public IReadOnlyList<string> Nodes { get; }

    /// <summary>Gets the target width in points.</summary>
    public double WidthPoints { get; }

    /// <summary>Gets the target height in points.</summary>
    public double HeightPoints { get; }
}
