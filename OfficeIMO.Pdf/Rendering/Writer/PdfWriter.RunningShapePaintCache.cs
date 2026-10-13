using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Measures immutable running-shape variants once during a document write.</summary>
    private sealed class RunningShapePaintCache {
        private readonly Dictionary<PdfHeaderFooterShape, RunningShapePaint> _entries = new();

        internal RunningShapePaint Get(PdfHeaderFooterShape variant, OfficeShape shape) {
            if (_entries.TryGetValue(variant, out var cached)) return cached;
            var report = new PdfConversionReport();
            var options = new PdfOptions().ReportDiagnosticsTo(report, "OfficeIMO.Pdf");
            var bounds = new List<RunningPaintBounds>();
            if (TryGetHeaderFooterShapeBounds(shape, 0, 0, options, "Running shape", out double x, out double y, out double width, out double height))
                bounds.Add(new("shape", x, y, width, height));
            foreach (OfficeShape marker in CreateLineMarkerShapes(shape)) {
                double bottomY = shape.Height - marker.Height;
                if (TryGetShapePaintBounds(marker, 0, bottomY, options, "Running shape", "marker", out x, out y, out width, out height))
                    bounds.Add(new("marker", x, y, width, height));
            }
            cached = new(bounds, report.Warnings.Select(w => w.Message).ToArray());
            _entries.Add(variant, cached);
            return cached;
        }
    }

    private sealed record RunningShapePaint(IReadOnlyList<RunningPaintBounds> Bounds, IReadOnlyList<string> UnmeasuredMessages);
    private readonly record struct RunningPaintBounds(string Kind, double X, double Y, double Width, double Height);
}
