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
    Cycle,

    /// <summary>Vertically ordered list nodes.</summary>
    List,

    /// <summary>Balanced row-and-column matrix nodes.</summary>
    Matrix,

    /// <summary>Stacked pyramid levels ordered from apex to base.</summary>
    Pyramid,

    /// <summary>Central concept with surrounding related nodes.</summary>
    Relationship
}

/// <summary>Representable visual styling for a semantic Office diagram.</summary>
public sealed class OfficeDiagramStyle {
    /// <summary>Creates a diagram style used by shared image and PDF renderers.</summary>
    public OfficeDiagramStyle(string fontFamily,
        IEnumerable<OfficeColor> nodeColors, OfficeColor nodeTextColor,
        OfficeColor nodeOutlineColor, OfficeColor connectorColor) {
        if (string.IsNullOrWhiteSpace(fontFamily)) {
            throw new ArgumentException("A diagram font family is required.",
                nameof(fontFamily));
        }
        if (nodeColors == null) throw new ArgumentNullException(nameof(nodeColors));
        var colors = new List<OfficeColor>(nodeColors);
        if (colors.Count == 0) {
            throw new ArgumentException("At least one diagram node color is required.",
                nameof(nodeColors));
        }
        FontFamily = fontFamily.Trim();
        NodeColors = new ReadOnlyCollection<OfficeColor>(colors);
        NodeTextColor = nodeTextColor;
        NodeOutlineColor = nodeOutlineColor;
        ConnectorColor = connectorColor;
    }

    /// <summary>Gets the node-label font family.</summary>
    public string FontFamily { get; }

    /// <summary>Gets node fill colors, repeated when there are more nodes.</summary>
    public IReadOnlyList<OfficeColor> NodeColors { get; }

    /// <summary>Gets the node-label color.</summary>
    public OfficeColor NodeTextColor { get; }

    /// <summary>Gets the node-outline color.</summary>
    public OfficeColor NodeOutlineColor { get; }

    /// <summary>Gets the connector color.</summary>
    public OfficeColor ConnectorColor { get; }
}

/// <summary>Dependency-free semantic diagram data for static rendering and export.</summary>
public sealed class OfficeDiagramSnapshot {
    /// <summary>Creates a semantic diagram snapshot.</summary>
    public OfficeDiagramSnapshot(string? name, OfficeDiagramKind kind,
        IEnumerable<string> nodes, double widthPoints,
        double heightPoints, OfficeDiagramStyle? style = null) {
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
        Style = style;
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

    /// <summary>Gets optional visual styling projected from the source diagram.</summary>
    public OfficeDiagramStyle? Style { get; }
}
