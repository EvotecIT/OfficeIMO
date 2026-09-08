using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Chooses the progression axis of complete visual column strips.</summary>
internal static class PdfTableColumnGeometry {
    internal static bool HasHorizontalProgression<T>(IReadOnlyList<T> columns,
        Func<T, PdfLogicalVisualBounds?> getBounds, bool fallback, CancellationToken cancellationToken = default) {
        if (columns.Count < 2) return fallback;
        double minX = double.MaxValue, maxX = double.MinValue, minY = double.MaxValue, maxY = double.MinValue;
        foreach (T column in columns) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfLogicalVisualBounds? bounds = getBounds(column);
            if (bounds is null) return fallback;
            double x = (bounds.Left + bounds.Right) / 2D, y = (bounds.Top + bounds.Bottom) / 2D;
            minX = Math.Min(minX, x); maxX = Math.Max(maxX, x);
            minY = Math.Min(minY, y); maxY = Math.Max(maxY, y);
        }
        double dx = maxX - minX, dy = maxY - minY;
        return Math.Max(dx, dy) <= 0.001D ? fallback : dx >= dy;
    }
}
