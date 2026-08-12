using System.Collections.Generic;

namespace OfficeIMO.ChartForgeX;

/// <summary>Describes one ChartForgeX artifact region in Office point coordinates.</summary>
public sealed class OfficeVisualRegion {
    internal OfficeVisualRegion(string id, string kind, string label, string? alternativeText, string? href,
        double? left, double? top, double? width, double? height, IReadOnlyDictionary<string, string> metadata) {
        Id = id;
        Kind = kind;
        Label = label;
        AlternativeText = alternativeText;
        Href = href;
        Left = left;
        Top = top;
        Width = width;
        Height = height;
        Metadata = metadata;
    }

    /// <summary>Gets the stable region identifier.</summary>
    public string Id { get; }

    /// <summary>Gets the product-neutral region kind.</summary>
    public string Kind { get; }

    /// <summary>Gets the display label.</summary>
    public string Label { get; }

    /// <summary>Gets an optional accessible text alternative.</summary>
    public string? AlternativeText { get; }

    /// <summary>Gets an optional region navigation target.</summary>
    public string? Href { get; }

    /// <summary>Gets the optional left coordinate in points.</summary>
    public double? Left { get; }

    /// <summary>Gets the optional top coordinate in points.</summary>
    public double? Top { get; }

    /// <summary>Gets the optional width in points.</summary>
    public double? Width { get; }

    /// <summary>Gets the optional height in points.</summary>
    public double? Height { get; }

    /// <summary>Gets region metadata copied from the source artifact.</summary>
    public IReadOnlyDictionary<string, string> Metadata { get; }
}
