using System;

namespace OfficeIMO.ChartForgeX;

/// <summary>Configures native editable Visio projection of a ChartForgeX visual artifact.</summary>
public sealed class OfficeVisioVisualOptions {
    /// <summary>Overrides graph styling. When null, card, border, and text colors follow the source theme with portable Arial text.</summary>
    public OfficeIMO.Visio.VisioStyleTheme? NativeTheme { get; set; }

    private OfficeVisioVisualLayoutMode _layoutMode = OfficeVisioVisualLayoutMode.Auto;

    /// <summary>Chooses prepared geometry or native reflow. Auto preserves complete topology bounds.</summary>
    public OfficeVisioVisualLayoutMode LayoutMode {
        get => _layoutMode;
        set {
            if (!Enum.IsDefined(typeof(OfficeVisioVisualLayoutMode), value)) throw new ArgumentOutOfRangeException(nameof(value));
            _layoutMode = value;
        }
    }

    /// <summary>Rejects conversion if any warning indicates semantic or presentation loss.</summary>
    public bool RequireLossless { get; set; }

    /// <summary>Specific diagnostic categories that must fail conversion even when other losses are accepted.</summary>
    public System.Collections.Generic.ISet<OfficeVisioVisualDiagnosticCode> RejectedDiagnostics { get; } = new System.Collections.Generic.HashSet<OfficeVisioVisualDiagnosticCode>();

    internal OfficeVisioVisualOptions ForPage(string name) {
        var copy = (OfficeVisioVisualOptions)MemberwiseClone();
        copy.PageName = name;
        return copy;
    }

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

    /// <summary>Gets or sets the pixel density used for preserved geometry and natural page sizing.</summary>
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

/// <summary>Controls native diagram layout during conversion.</summary>
public enum OfficeVisioVisualLayoutMode {
    /// <summary>Preserve complete topology bounds; otherwise use native reflow.</summary>
    Auto,
    /// <summary>Require topology bounds and preserve their geometry. Missing bounds fail closed.</summary>
    Preserve,
    /// <summary>Use the native Visio layout and report normalization.</summary>
    Reflow
}
