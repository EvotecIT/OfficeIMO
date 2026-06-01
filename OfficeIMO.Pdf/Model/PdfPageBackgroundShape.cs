using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable page background shape rendered behind page content.
/// </summary>
public sealed class PdfPageBackgroundShape {
    private OfficeShape _shape;
    private double _x;
    private double _y;

    /// <summary>Creates a page background shape at an absolute page position in points.</summary>
    public PdfPageBackgroundShape(OfficeShape shape, double x, double y) {
        Guard.NotNull(shape, nameof(shape));
        ValidateFinite(x, nameof(x));
        ValidateFinite(y, nameof(y));
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
    public double X {
        get => _x;
        set {
            ValidateFinite(value, nameof(X));
            _x = value;
        }
    }

    /// <summary>Bottom position in points from the page bottom edge. Negative values are allowed for bleed-style decoration.</summary>
    public double Y {
        get => _y;
        set {
            ValidateFinite(value, nameof(Y));
            _y = value;
        }
    }

    /// <summary>Creates a rectangle background shape.</summary>
    public static PdfPageBackgroundShape Rectangle(double x, double y, double width, double height, PdfColor? fill = null, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) {
        var shape = OfficeShape.Rectangle(width, height);
        ApplyColors(shape, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient);
        return new PdfPageBackgroundShape(shape, x, y);
    }

    /// <summary>Creates a rounded rectangle background shape.</summary>
    public static PdfPageBackgroundShape RoundedRectangle(double x, double y, double width, double height, double cornerRadius, PdfColor? fill = null, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) {
        var shape = OfficeShape.RoundedRectangle(width, height, cornerRadius);
        ApplyColors(shape, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient);
        return new PdfPageBackgroundShape(shape, x, y);
    }

    /// <summary>Creates an ellipse background shape.</summary>
    public static PdfPageBackgroundShape Ellipse(double x, double y, double width, double height, PdfColor? fill = null, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) {
        var shape = OfficeShape.Ellipse(width, height);
        ApplyColors(shape, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient);
        return new PdfPageBackgroundShape(shape, x, y);
    }

    /// <summary>Creates a full-width band anchored to the top of a page.</summary>
    public static PdfPageBackgroundShape TopBand(double pageWidth, double pageHeight, double height, PdfColor? fill = null, double insetX = 0D, double offsetY = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) {
        ValidateBandPage(pageWidth, pageHeight);
        Guard.Positive(height, nameof(height));
        Guard.NonNegative(insetX, nameof(insetX));
        Guard.NonNegative(offsetY, nameof(offsetY));
        Guard.NonNegative(cornerRadius, nameof(cornerRadius));
        double width = pageWidth - insetX * 2D;
        double y = pageHeight - offsetY - height;
        ValidateBandBounds(width, height, insetX, y, pageWidth, pageHeight);
        return cornerRadius > 0D
            ? RoundedRectangle(insetX, y, width, height, cornerRadius, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient)
            : Rectangle(insetX, y, width, height, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient);
    }

    /// <summary>Creates a full-width band anchored to the bottom of a page.</summary>
    public static PdfPageBackgroundShape BottomBand(double pageWidth, double pageHeight, double height, PdfColor? fill = null, double insetX = 0D, double offsetY = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) {
        ValidateBandPage(pageWidth, pageHeight);
        Guard.Positive(height, nameof(height));
        Guard.NonNegative(insetX, nameof(insetX));
        Guard.NonNegative(offsetY, nameof(offsetY));
        Guard.NonNegative(cornerRadius, nameof(cornerRadius));
        double width = pageWidth - insetX * 2D;
        ValidateBandBounds(width, height, insetX, offsetY, pageWidth, pageHeight);
        return cornerRadius > 0D
            ? RoundedRectangle(insetX, offsetY, width, height, cornerRadius, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient)
            : Rectangle(insetX, offsetY, width, height, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient);
    }

    /// <summary>Creates a full-height band anchored to the left side of a page.</summary>
    public static PdfPageBackgroundShape LeftBand(double pageWidth, double pageHeight, double width, PdfColor? fill = null, double insetY = 0D, double offsetX = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) {
        ValidateBandPage(pageWidth, pageHeight);
        Guard.Positive(width, nameof(width));
        Guard.NonNegative(insetY, nameof(insetY));
        Guard.NonNegative(offsetX, nameof(offsetX));
        Guard.NonNegative(cornerRadius, nameof(cornerRadius));
        double height = pageHeight - insetY * 2D;
        ValidateBandBounds(width, height, offsetX, insetY, pageWidth, pageHeight);
        return cornerRadius > 0D
            ? RoundedRectangle(offsetX, insetY, width, height, cornerRadius, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient)
            : Rectangle(offsetX, insetY, width, height, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient);
    }

    /// <summary>Creates a full-height band anchored to the right side of a page.</summary>
    public static PdfPageBackgroundShape RightBand(double pageWidth, double pageHeight, double width, PdfColor? fill = null, double insetY = 0D, double offsetX = 0D, double cornerRadius = 0D, PdfColor? stroke = null, double strokeWidth = 0D, double? fillOpacity = null, double? strokeOpacity = null, OfficeLinearGradient? fillGradient = null) {
        ValidateBandPage(pageWidth, pageHeight);
        Guard.Positive(width, nameof(width));
        Guard.NonNegative(insetY, nameof(insetY));
        Guard.NonNegative(offsetX, nameof(offsetX));
        Guard.NonNegative(cornerRadius, nameof(cornerRadius));
        double height = pageHeight - insetY * 2D;
        double x = pageWidth - offsetX - width;
        ValidateBandBounds(width, height, x, insetY, pageWidth, pageHeight);
        return cornerRadius > 0D
            ? RoundedRectangle(x, insetY, width, height, cornerRadius, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient)
            : Rectangle(x, insetY, width, height, fill, stroke, strokeWidth, fillOpacity, strokeOpacity, fillGradient);
    }

    /// <summary>Creates a deep copy of this page background shape.</summary>
    public PdfPageBackgroundShape Clone() => new PdfPageBackgroundShape(_shape, X, Y);

    private static void ApplyColors(OfficeShape shape, PdfColor? fill, PdfColor? stroke, double strokeWidth, double? fillOpacity, double? strokeOpacity, OfficeLinearGradient? fillGradient) {
        if (fill.HasValue) {
            shape.FillColor = fill.Value.ToOfficeColor();
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
            shape.StrokeColor = stroke.Value.ToOfficeColor();
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
