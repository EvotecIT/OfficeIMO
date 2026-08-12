using System;

namespace OfficeIMO.ChartForgeX;

/// <summary>Configures native editable Visio projection of a ChartForgeX visual artifact.</summary>
public sealed class OfficeVisioVisualOptions {
    private string _pageName = "Visual Artifact";
    private double _pixelsPerInch = 96D;

    /// <summary>Gets or sets the Visio page name.</summary>
    public string PageName {
        get => _pageName;
        set => _pageName = string.IsNullOrWhiteSpace(value) ? throw new ArgumentException("Page name cannot be null or whitespace.", nameof(value)) : value;
    }

    /// <summary>Gets or sets whether a non-empty artifact title is added as an editable Visio title.</summary>
    public bool IncludeTitle { get; set; } = true;

    /// <summary>Gets or sets whether topology and flow groups become editable Visio containers.</summary>
    public bool IncludeGroups { get; set; } = true;

    /// <summary>Gets or sets whether product-neutral metadata and details are written as Visio Shape Data.</summary>
    public bool IncludeShapeData { get; set; } = true;

    /// <summary>Gets or sets whether safe CFX hyperlinks are attached to native Visio shapes and connectors.</summary>
    public bool IncludeHyperlinks { get; set; } = true;

    /// <summary>Gets or sets whether the CFX natural pixel size is used as the minimum Visio page size.</summary>
    /// <remarks>The default is false so native Visio builders size the page to their editable content.</remarks>
    public bool UseNaturalPageSize { get; set; }

    /// <summary>Gets or sets the pixel density used when <see cref="UseNaturalPageSize"/> is enabled.</summary>
    public double PixelsPerInch {
        get => _pixelsPerInch;
        set {
            if (value <= 0D || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new ArgumentOutOfRangeException(nameof(value), value, "Pixels per inch must be positive and finite.");
            }
            _pixelsPerInch = value;
        }
    }
}
