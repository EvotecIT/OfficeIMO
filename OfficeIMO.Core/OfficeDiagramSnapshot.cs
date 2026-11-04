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
        IEnumerable<string> nodes, OfficeLength width,
        OfficeLength height, OfficeDiagramStyle? style = null) {
        if (nodes == null) throw new ArgumentNullException(nameof(nodes));
        if (width <= OfficeLength.Zero) {
            throw new ArgumentOutOfRangeException(nameof(width));
        }
        if (height <= OfficeLength.Zero) {
            throw new ArgumentOutOfRangeException(nameof(height));
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
        Width = width;
        Height = height;
        Style = style;
    }

    /// <summary>Gets the optional source diagram name.</summary>
    public string? Name { get; }

    /// <summary>Gets the semantic layout kind.</summary>
    public OfficeDiagramKind Kind { get; }

    /// <summary>Gets node labels in semantic order.</summary>
    public IReadOnlyList<string> Nodes { get; }

    /// <summary>Gets the physical target width.</summary>
    public OfficeLength Width { get; }

    /// <summary>Gets the physical target height.</summary>
    public OfficeLength Height { get; }

    /// <summary>Gets optional visual styling projected from the source diagram.</summary>
    public OfficeDiagramStyle? Style { get; }
}
