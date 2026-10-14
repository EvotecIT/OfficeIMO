using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal static partial class VisioForeignImage {
    internal static bool TryGetProjection(VisioShape shape, VisioPage page, double scale,
        ICollection<OfficeImageExportDiagnostic>? diagnostics, string? source, out OfficeImageProjection projection) {
        projection = default;
        try {
            double x = ReadCell(shape, "ImgOffsetX", 0, diagnostics, source), y = ReadCell(shape, "ImgOffsetY", 0, diagnostics, source);
            double width = ReadCell(shape, "ImgWidth", shape.Width, diagnostics, source), height = ReadCell(shape, "ImgHeight", shape.Height, diagnostics, source);
            if (!Finite(width) || !Finite(height) || width <= 0 || height <= 0 || !Finite(shape.Width) || !Finite(shape.Height) || shape.Width < 0 || shape.Height < 0)
                throw new InvalidDataException("Foreign image dimensions must be finite and positive.");
            double left = Math.Max(0, x), right = Math.Min(shape.Width, x + width);
            double bottom = Math.Max(0, y), top = Math.Min(shape.Height, y + height);
            if (right <= left || top <= bottom) return false; // A fully cropped image is intentionally invisible.
            var crop = OfficeImageSourceCrop.FromStrictFractions((left - x) / width, (y + height - top) / height,
                (x + width - right) / width, (bottom - y) / height);
            OfficeTransform transform = PageTransform(shape, page.Height, diagnostics, source);
            OfficePoint center = transform.TransformPoint(new OfficePoint((left + right) / 2, (bottom + top) / 2));
            OfficePoint a = transform.TransformPoint(new OfficePoint(left, top));
            OfficePoint b = transform.TransformPoint(new OfficePoint(right, top));
            OfficePoint c = transform.TransformPoint(new OfficePoint(left, bottom));
            double angle = Math.Atan2(b.Y - a.Y, b.X - a.X) * 180 / Math.PI;
            bool mirrored = (b.X - a.X) * (c.Y - a.Y) - (b.Y - a.Y) * (c.X - a.X) < 0;
            double w = (right - left) * scale, h = (top - bottom) * scale;
            projection = new OfficeImageProjection(new OfficeImagePlacement(center.X * scale - w / 2, center.Y * scale - h / 2, w, h),
                crop, angle, center.X * scale, center.Y * scale, flipVertical: mirrored);
            return true;
        } catch (Exception exception) when (exception is ArgumentException || exception is InvalidDataException || exception is OverflowException) {
            diagnostics?.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                "VISIO_FOREIGN_PLACEMENT", exception.Message, Location(shape, source), OfficeConversionLossKind.Omission));
            return false;
        }
    }

    private static OfficeTransform PageTransform(VisioShape shape, double pageHeight,
        ICollection<OfficeImageExportDiagnostic>? diagnostics, string? source) {
        OfficeTransform transform = VisioNativeShapeTransform.Create(shape, diagnostics, source).Matrix;
        return transform.Then(OfficeTransform.Scale(1, -1)).Then(OfficeTransform.Translate(0, pageHeight));
    }

    private static double ReadCell(VisioShape shape, string name, double fallback,
        ICollection<OfficeImageExportDiagnostic>? diagnostics, string? source) {
        XElement? Find(VisioShape owner) => owner.PreservedCellElements.FirstOrDefault(cell => (string?)cell.Attribute("N") == name);
        XElement? cell = Find(shape);
        if (cell == null || (string?)cell.Attribute("F") == "Inh")
            cell = MasterShape(shape) is VisioShape master ? Find(master) ?? cell : cell;
        if (cell == null) return fallback;
        string? formula = (string?)cell.Attribute("F");
        if (!string.IsNullOrWhiteSpace(formula) && formula != "Inh") {
            if (VisioShapeGeometry.TryParseCellLiteral(formula, shape, out double evaluated)) return evaluated;
            diagnostics?.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                "VISIO_FOREIGN_CACHED_FORMULA", "Image placement uses a cached value for unsupported formula in " + name + ".",
                Location(shape, source), OfficeConversionLossKind.Approximation));
        }
        if (VisioShapeGeometry.TryParseCellLiteral((string?)cell.Attribute("V"), shape, out double value)) return value;
        throw new InvalidDataException("Foreign image cell " + name + " has no usable finite value.");
    }
    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
