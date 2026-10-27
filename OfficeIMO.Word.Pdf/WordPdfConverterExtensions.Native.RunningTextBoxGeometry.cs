using System.Globalization;
using A = DocumentFormat.OpenXml.Drawing;
using WP14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Resolves persisted page anchors and fractional relative extents without modifying source XML.</summary>
    private static bool TryGetNativeRunningTextBoxFrame(WordTextBox box, WordSection section,
        WordPdfExportOptions? options, PdfCore.PdfPageSize page, out double x, out double y, out double width, out double height,
        out bool behindDocument) {
        x = y = width = height = 0D;
        behindDocument = false;
        var properties = box.DrawingShapeProperties;
        if (properties?.GetFirstChild<A.CustomGeometry>() != null ||
            !string.Equals(properties?.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText,
                "rect", StringComparison.Ordinal)) return false;
        if (!WordDrawingLayoutReader.TryRead(box.Drawing, out WordDrawingLayoutInfo layout) ||
            layout.Placement != WordDrawingPlacementKind.Anchored || layout.Wrap != WordDrawingWrapKind.Square ||
            layout.HorizontalRelativeFrom != "page" || layout.VerticalRelativeFrom != "page" ||
            box.DrawingShapeProperties?.GetFirstChild<A.Transform2D>() is { } transform &&
                (transform.Rotation?.Value is > 0 or < 0 || transform.HorizontalFlip?.Value == true || transform.VerticalFlip?.Value == true)) return false;
        behindDocument = layout.BehindDocument;
        PdfCore.PdfPageMargins margins = GetNativeMargins(section, options, (0D, 0D));
        width = layout.WidthPoints;
        height = layout.HeightPoints;
        WP14.RelativeWidth? relativeWidth = box.Anchor?.GetFirstChild<WP14.RelativeWidth>();
        WP14.RelativeHeight? relativeHeight = box.Anchor?.GetFirstChild<WP14.RelativeHeight>();
        if (!TryGetNativeRunningRelativeExtent(relativeWidth?.PercentageWidth?.Text, relativeWidth?.ObjectId?.InnerText,
                page.Width.ToPoints(), margins.Left.ToPoints(), margins.Right.ToPoints(), ref width) ||
            !TryGetNativeRunningRelativeExtent(relativeHeight?.PercentageHeight?.Text, relativeHeight?.RelativeFrom?.InnerText,
                page.Height.ToPoints(), margins.Top.ToPoints(), margins.Bottom.ToPoints(), ref height)) return false;
        if (!TryGetNativeGroupAxis("page", layout.HorizontalOffsetPoints, layout.HorizontalAlignment,
                page.Width.ToPoints(), 0D, 0D, width, false, out x) ||
            !TryGetNativeGroupAxis("page", layout.VerticalOffsetPoints, layout.VerticalAlignment,
                page.Height.ToPoints(), 0D, 0D, height, true, out y)) return false;
        return width > 0D && height > 0D && x >= 0D && y >= 0D &&
            x + width <= page.Width.ToPoints() + .001D && y + height <= page.Height.ToPoints() + .001D;
    }

    private static bool TryGetNativeRunningRelativeExtent(string? percentage, string? reference,
        double pageLength, double leadingMargin, double trailingMargin, ref double extent) {
        if (percentage == null) return true;
        if (!double.TryParse(percentage, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
            double.IsNaN(value) || double.IsInfinity(value) || value < 0D) return false;
        if (value == 0D) return true;
        if (reference != "page" && reference != "margin") return false;
        extent = (reference == "page" ? pageLength : pageLength - leadingMargin - trailingMargin) * value / 100000D;
        return !double.IsNaN(extent) && !double.IsInfinity(extent);
    }
}
