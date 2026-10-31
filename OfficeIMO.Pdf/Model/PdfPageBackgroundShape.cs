using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable page background shape rendered behind page content.
/// </summary>
public sealed class PdfPageBackgroundShape {
    private OfficeShape _shape;
    private global::OfficeIMO.OfficeLength _x;
    private global::OfficeIMO.OfficeLength _y;

    /// <summary>Creates a page background shape at an absolute page position in points.</summary>
    public PdfPageBackgroundShape(OfficeShape shape, global::OfficeIMO.OfficeLength x, global::OfficeIMO.OfficeLength y) {
        Guard.NotNull(shape, nameof(shape));
        ValidateFinite(x.ToPoints(), nameof(x));
        ValidateFinite(y.ToPoints(), nameof(y));
        _shape = shape.Clone();
        _x = x;
        _y = y;
    }

    /// <summary>Shape geometry and styling. Coordinates use the page bottom-left origin.</summary>
    public OfficeShape Shape {
        get => _shape.Clone();
        set {
            Guard.NotNull(value, nameof(Shape));
            _shape = value.Clone();
        }
    }

    /// <summary>Left position in points from the page left edge. Negative values are allowed for bleed-style decoration.</summary>
    public global::OfficeIMO.OfficeLength X {
        get => _x;
        set {
            ValidateFinite(value.ToPoints(), nameof(X));
            _x = value;
        }
    }

    /// <summary>Bottom position in points from the page bottom edge. Negative values are allowed for bleed-style decoration.</summary>
    public global::OfficeIMO.OfficeLength Y {
        get => _y;
        set {
            ValidateFinite(value.ToPoints(), nameof(Y));
            _y = value;
        }
    }

    /// <summary>
    /// When true, the shape is drawn on the first page only, such as a title band that would otherwise sit behind
    /// body text on continuation pages.
    /// </summary>
    public bool FirstPageOnly { get; set; }

    /// <summary>Creates a rectangle background shape.</summary>
    public static PdfPageBackgroundShape Rectangle(global::OfficeIMO.OfficeLength x, global::OfficeIMO.OfficeLength y, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, global::OfficeIMO.OfficeColor? fill = null, global::OfficeIMO.OfficeColor? stroke = null, global::OfficeIMO.OfficeLength strokeWidth = default, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) {
        var shape = OfficeShape.Rectangle(width.ToPoints(), height.ToPoints());
        ApplyColors(shape, fill, stroke, strokeWidth.ToPoints(), fillOpacity, strokeOpacity, fillGradient);
        return new PdfPageBackgroundShape(shape, x, y);
    }

    /// <summary>Creates a rounded rectangle background shape.</summary>
    public static PdfPageBackgroundShape RoundedRectangle(global::OfficeIMO.OfficeLength x, global::OfficeIMO.OfficeLength y, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, global::OfficeIMO.OfficeLength cornerRadius, global::OfficeIMO.OfficeColor? fill = null, global::OfficeIMO.OfficeColor? stroke = null, global::OfficeIMO.OfficeLength strokeWidth = default, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) {
        var shape = OfficeShape.RoundedRectangle(width.ToPoints(), height.ToPoints(), cornerRadius.ToPoints());
        ApplyColors(shape, fill, stroke, strokeWidth.ToPoints(), fillOpacity, strokeOpacity, fillGradient);
        return new PdfPageBackgroundShape(shape, x, y);
    }

    /// <summary>Creates an ellipse background shape.</summary>
    public static PdfPageBackgroundShape Ellipse(global::OfficeIMO.OfficeLength x, global::OfficeIMO.OfficeLength y, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, global::OfficeIMO.OfficeColor? fill = null, global::OfficeIMO.OfficeColor? stroke = null, global::OfficeIMO.OfficeLength strokeWidth = default, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) {
        var shape = OfficeShape.Ellipse(width.ToPoints(), height.ToPoints());
        ApplyColors(shape, fill, stroke, strokeWidth.ToPoints(), fillOpacity, strokeOpacity, fillGradient);
        return new PdfPageBackgroundShape(shape, x, y);
    }

    /// <summary>Creates a full-width band anchored to the top of a page.</summary>
    public static PdfPageBackgroundShape TopBand(global::OfficeIMO.OfficeLength pageWidth, global::OfficeIMO.OfficeLength pageHeight, global::OfficeIMO.OfficeLength height, global::OfficeIMO.OfficeColor? fill = null, global::OfficeIMO.OfficeLength insetX = default, global::OfficeIMO.OfficeLength offsetY = default, global::OfficeIMO.OfficeLength cornerRadius = default, global::OfficeIMO.OfficeColor? stroke = null, global::OfficeIMO.OfficeLength strokeWidth = default, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) {
        ValidateBandPage(pageWidth.ToPoints(), pageHeight.ToPoints());
        Guard.Positive(height.ToPoints(), nameof(height));
        Guard.NonNegative(insetX.ToPoints(), nameof(insetX));
        Guard.NonNegative(offsetY.ToPoints(), nameof(offsetY));
        Guard.NonNegative(cornerRadius.ToPoints(), nameof(cornerRadius));
        double width = pageWidth.ToPoints() - insetX.ToPoints() * 2D;
        double y = pageHeight.ToPoints() - offsetY.ToPoints() - height.ToPoints();
        ValidateBandBounds(width, height.ToPoints(), insetX.ToPoints(), y, pageWidth.ToPoints(), pageHeight.ToPoints());
        return cornerRadius.ToPoints() > 0D
            ? RoundedRectangle(insetX, global::OfficeIMO.OfficeLength.Points(y), global::OfficeIMO.OfficeLength.Points(width), height, cornerRadius, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient)
            : Rectangle(insetX, global::OfficeIMO.OfficeLength.Points(y), global::OfficeIMO.OfficeLength.Points(width), height, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient);
    }

    /// <summary>Creates a full-width band anchored to the bottom of a page.</summary>
    public static PdfPageBackgroundShape BottomBand(global::OfficeIMO.OfficeLength pageWidth, global::OfficeIMO.OfficeLength pageHeight, global::OfficeIMO.OfficeLength height, global::OfficeIMO.OfficeColor? fill = null, global::OfficeIMO.OfficeLength insetX = default, global::OfficeIMO.OfficeLength offsetY = default, global::OfficeIMO.OfficeLength cornerRadius = default, global::OfficeIMO.OfficeColor? stroke = null, global::OfficeIMO.OfficeLength strokeWidth = default, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) {
        ValidateBandPage(pageWidth.ToPoints(), pageHeight.ToPoints());
        Guard.Positive(height.ToPoints(), nameof(height));
        Guard.NonNegative(insetX.ToPoints(), nameof(insetX));
        Guard.NonNegative(offsetY.ToPoints(), nameof(offsetY));
        Guard.NonNegative(cornerRadius.ToPoints(), nameof(cornerRadius));
        double width = pageWidth.ToPoints() - insetX.ToPoints() * 2D;
        ValidateBandBounds(width, height.ToPoints(), insetX.ToPoints(), offsetY.ToPoints(), pageWidth.ToPoints(), pageHeight.ToPoints());
        return cornerRadius.ToPoints() > 0D
            ? RoundedRectangle(insetX, offsetY, global::OfficeIMO.OfficeLength.Points(width), height, cornerRadius, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient)
            : Rectangle(insetX, offsetY, global::OfficeIMO.OfficeLength.Points(width), height, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient);
    }

    /// <summary>Creates a full-height band anchored to the left side of a page.</summary>
    public static PdfPageBackgroundShape LeftBand(global::OfficeIMO.OfficeLength pageWidth, global::OfficeIMO.OfficeLength pageHeight, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeColor? fill = null, global::OfficeIMO.OfficeLength insetY = default, global::OfficeIMO.OfficeLength offsetX = default, global::OfficeIMO.OfficeLength cornerRadius = default, global::OfficeIMO.OfficeColor? stroke = null, global::OfficeIMO.OfficeLength strokeWidth = default, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) {
        ValidateBandPage(pageWidth.ToPoints(), pageHeight.ToPoints());
        Guard.Positive(width.ToPoints(), nameof(width));
        Guard.NonNegative(insetY.ToPoints(), nameof(insetY));
        Guard.NonNegative(offsetX.ToPoints(), nameof(offsetX));
        Guard.NonNegative(cornerRadius.ToPoints(), nameof(cornerRadius));
        double height = pageHeight.ToPoints() - insetY.ToPoints() * 2D;
        ValidateBandBounds(width.ToPoints(), height, offsetX.ToPoints(), insetY.ToPoints(), pageWidth.ToPoints(), pageHeight.ToPoints());
        return cornerRadius.ToPoints() > 0D
            ? RoundedRectangle(offsetX, insetY, width, global::OfficeIMO.OfficeLength.Points(height), cornerRadius, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient)
            : Rectangle(offsetX, insetY, width, global::OfficeIMO.OfficeLength.Points(height), fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient);
    }

    /// <summary>Creates a full-height band anchored to the right side of a page.</summary>
    public static PdfPageBackgroundShape RightBand(global::OfficeIMO.OfficeLength pageWidth, global::OfficeIMO.OfficeLength pageHeight, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeColor? fill = null, global::OfficeIMO.OfficeLength insetY = default, global::OfficeIMO.OfficeLength offsetX = default, global::OfficeIMO.OfficeLength cornerRadius = default, global::OfficeIMO.OfficeColor? stroke = null, global::OfficeIMO.OfficeLength strokeWidth = default, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) {
        ValidateBandPage(pageWidth.ToPoints(), pageHeight.ToPoints());
        Guard.Positive(width.ToPoints(), nameof(width));
        Guard.NonNegative(insetY.ToPoints(), nameof(insetY));
        Guard.NonNegative(offsetX.ToPoints(), nameof(offsetX));
        Guard.NonNegative(cornerRadius.ToPoints(), nameof(cornerRadius));
        double height = pageHeight.ToPoints() - insetY.ToPoints() * 2D;
        double x = pageWidth.ToPoints() - offsetX.ToPoints() - width.ToPoints();
        ValidateBandBounds(width.ToPoints(), height, x, insetY.ToPoints(), pageWidth.ToPoints(), pageHeight.ToPoints());
        return cornerRadius.ToPoints() > 0D
            ? RoundedRectangle(global::OfficeIMO.OfficeLength.Points(x), insetY, width, global::OfficeIMO.OfficeLength.Points(height), cornerRadius, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient)
            : Rectangle(global::OfficeIMO.OfficeLength.Points(x), insetY, width, global::OfficeIMO.OfficeLength.Points(height), fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient);
    }

    /// <summary>Creates a deep copy of this page background shape.</summary>
    public PdfPageBackgroundShape Clone() => new PdfPageBackgroundShape(_shape, X, Y) { FirstPageOnly = FirstPageOnly };

    private static void ApplyColors(OfficeShape shape, global::OfficeIMO.OfficeColor? fill, global::OfficeIMO.OfficeColor? stroke, double strokeWidth, double? fillOpacity, double? strokeOpacity, OfficeLinearGradient? fillGradient) {
        if (fill.HasValue) {
            shape.FillColor = fill.Value;
        }

        if (fillGradient != null) {
            shape.FillGradient = fillGradient.Clone();
        }

        if (fillOpacity.HasValue) {
            ValidateOpacity(fillOpacity.Value, nameof(fillOpacity));
            shape.FillOpacity = fillOpacity.Value;
        }

        if (stroke.HasValue) {
            Guard.Positive(strokeWidth, nameof(strokeWidth));
            shape.StrokeColor = stroke.Value;
            shape.StrokeWidth = strokeWidth;
        } else {
            Guard.NonNegative(strokeWidth, nameof(strokeWidth));
            shape.StrokeColor = null;
            shape.StrokeWidth = 0D;
        }

        if (strokeOpacity.HasValue) {
            ValidateOpacity(strokeOpacity.Value, nameof(strokeOpacity));
            shape.StrokeOpacity = strokeOpacity.Value;
        }
    }

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentOutOfRangeException(paramName, value, "PDF page background shape coordinates must be finite numbers.");
        }
    }

    private static void ValidateBandPage(double pageWidth, double pageHeight) {
        Guard.Positive(pageWidth, nameof(pageWidth));
        Guard.Positive(pageHeight, nameof(pageHeight));
    }

    private static void ValidateBandBounds(double width, double height, double x, double y, double pageWidth, double pageHeight) {
        if (width <= 0D || height <= 0D || x < 0D || y < 0D || x + width > pageWidth || y + height > pageHeight) {
            throw new System.ArgumentException("PDF page background band must fit inside the page bounds.");
        }
    }

    private static void ValidateOpacity(double value, string paramName) {
        if (value < 0D || value > 1D || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentOutOfRangeException(paramName, value, "PDF page background shape opacity must be a finite number between 0 and 1.");
        }
    }
}
