using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote;

/// <summary>Controls dependency-free OneNote page layout, rendering, and image export.</summary>
public class OneNotePageRenderingOptions : OfficeImageExportOptions {
    /// <inheritdoc />
    public override double LogicalUnitsPerInch => 72D;

    /// <summary>Whether the page title is rendered.</summary>
    public bool IncludeTitle { get; set; } = true;

    /// <summary>Whether embedded pictures, including printout backgrounds, are rendered.</summary>
    public bool IncludeImages { get; set; } = true;

    /// <summary>Whether native ink strokes are rendered.</summary>
    public bool IncludeInk { get; set; } = true;

    /// <summary>Whether structured mathematical expressions are typeset.</summary>
    public bool IncludeMath { get; set; } = true;

    /// <summary>Whether attachment and recording placeholders are rendered.</summary>
    public bool IncludeAttachmentPlaceholders { get; set; } = true;

    /// <summary>Maximum bytes materialized from any single lazy image payload.</summary>
    public long MaxImageBytes { get; set; } = 64L * 1024L * 1024L;

    /// <summary>Minimum width used for automatically sized pages, in points.</summary>
    public double AutomaticPageWidthPoints { get; set; } = 612D;

    /// <summary>Minimum height used for automatically sized pages, in points.</summary>
    public double AutomaticPageHeightPoints { get; set; } = 792D;

    /// <summary>Extra space retained beyond inferred content bounds, in points.</summary>
    public double AutomaticPagePaddingPoints { get; set; } = 36D;

    /// <summary>Default body font used when a OneNote run does not name one.</summary>
    public OfficeFontInfo DefaultFont { get; set; } = new OfficeFontInfo("Calibri", 11D);

    /// <summary>Reusable ink-rendering settings.</summary>
    public OfficeInkRenderOptions Ink { get; set; } = new OfficeInkRenderOptions();

    /// <summary>Reusable mathematical-rendering settings.</summary>
    public OfficeMathRenderOptions Math { get; set; } = new OfficeMathRenderOptions();

    /// <summary>Creates a detached copy.</summary>
    public OneNotePageRenderingOptions Clone() => CopyTo(new OneNotePageRenderingOptions());

    internal T CopyTo<T>(T clone) where T : OneNotePageRenderingOptions {
        CopyImageExportOptionsTo(clone);
        clone.IncludeTitle = IncludeTitle;
        clone.IncludeImages = IncludeImages;
        clone.IncludeInk = IncludeInk;
        clone.IncludeMath = IncludeMath;
        clone.IncludeAttachmentPlaceholders = IncludeAttachmentPlaceholders;
        clone.MaxImageBytes = MaxImageBytes;
        clone.AutomaticPageWidthPoints = AutomaticPageWidthPoints;
        clone.AutomaticPageHeightPoints = AutomaticPageHeightPoints;
        clone.AutomaticPagePaddingPoints = AutomaticPagePaddingPoints;
        clone.DefaultFont = DefaultFont;
        clone.Ink = Ink?.Clone() ?? new OfficeInkRenderOptions();
        clone.Math = Math?.Clone() ?? new OfficeMathRenderOptions();
        return clone;
    }

    internal void Validate() {
        ValidateImageExportOptions();
        if (MaxImageBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaxImageBytes));
        ValidatePositive(AutomaticPageWidthPoints, nameof(AutomaticPageWidthPoints));
        ValidatePositive(AutomaticPageHeightPoints, nameof(AutomaticPageHeightPoints));
        if (double.IsNaN(AutomaticPagePaddingPoints) || double.IsInfinity(AutomaticPagePaddingPoints) || AutomaticPagePaddingPoints < 0D) {
            throw new ArgumentOutOfRangeException(nameof(AutomaticPagePaddingPoints));
        }
        if (DefaultFont.Size <= 0D || double.IsNaN(DefaultFont.Size) || double.IsInfinity(DefaultFont.Size)) {
            throw new ArgumentOutOfRangeException(nameof(DefaultFont));
        }
        if (Ink == null) throw new InvalidOperationException("Ink rendering options cannot be null.");
        if (Math == null) throw new InvalidOperationException("Math rendering options cannot be null.");
    }

    private static void ValidatePositive(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) throw new ArgumentOutOfRangeException(name);
    }
}
