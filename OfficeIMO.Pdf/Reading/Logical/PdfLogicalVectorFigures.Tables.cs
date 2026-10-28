using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfLogicalVectorFigureAnalysis {
    /// <summary>
    /// Recognizes adjacent closed cells, rather than the open axes and separated
    /// marks of a chart. This rejects figure classification; it does not claim
    /// that the native table detector reconstructed the whole grid.
    /// </summary>
    private static bool IsRuledCellGrid(Box[] members, IReadOnlyList<Box> tableEvidence, CancellationToken cancellationToken) {
        const double tolerance = 0.75D;
        double left = members.Min(static box => box.Left), right = members.Max(static box => box.Right);
        double top = members.Min(static box => box.Top), bottom = members.Max(static box => box.Bottom);
        // Partial text-table ownership may have removed whole cells from members.
        // Retain their physical grid evidence inside this figure's region.
        Box[] region = tableEvidence.Where(box => box.Left >= left - tolerance && box.Right <= right + tolerance &&
            box.Top >= top - tolerance && box.Bottom <= bottom + tolerance).ToArray();
        Box[] rules = region.Where(static box => box.Stroked &&
            (box.Right - box.Left <= tolerance || box.Bottom - box.Top <= tolerance)).ToArray();
        var cells = new List<Box>();
        foreach (Box box in region) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!box.Rectangle || box.Right - box.Left < 12D || box.Bottom - box.Top < 6D) continue;
            if (box.Stroked || HasEdge(box, horizontal: true, box.Top) && HasEdge(box, horizontal: true, box.Bottom) &&
                HasEdge(box, horizontal: false, box.Left) && HasEdge(box, horizontal: false, box.Right)) cells.Add(box);
        }
        if (cells.Count < 3) return false;
        int[] parent = Enumerable.Range(0, cells.Count).ToArray();
        int[] size = Enumerable.Repeat(1, cells.Count).ToArray();
        var acrossColumns = new bool[cells.Count];
        var acrossRows = new bool[cells.Count];
        for (int first = 0; first < cells.Count; first++) {
            for (int second = first + 1; second < cells.Count; second++) {
                cancellationToken.ThrowIfCancellationRequested();
                Box a = cells[first], b = cells[second];
                bool columns = Near(a.Top, b.Top) && Near(a.Bottom, b.Bottom) &&
                    (Near(a.Right, b.Left) || Near(b.Right, a.Left));
                bool rows = Near(a.Left, b.Left) && Near(a.Right, b.Right) &&
                    (Near(a.Bottom, b.Top) || Near(b.Bottom, a.Top));
                if (!columns && !rows) continue;
                int firstRoot = Find(parent, first), secondRoot = Find(parent, second);
                if (firstRoot != secondRoot) {
                    parent[secondRoot] = firstRoot;
                    size[firstRoot] += size[secondRoot];
                    acrossColumns[firstRoot] |= acrossColumns[secondRoot];
                    acrossRows[firstRoot] |= acrossRows[secondRoot];
                }
                acrossColumns[firstRoot] |= columns;
                acrossRows[firstRoot] |= rows;
                // A single chain can be a stacked bar. Require a second grid
                // dimension, supplied either by cells or visible inner rulings.
                if (size[firstRoot] >= 3 && (acrossColumns[firstRoot] && acrossRows[firstRoot] || HasInnerRuling(firstRoot))) return true;
            }
        }
        return false;

        bool Near(double first, double second) => Math.Abs(first - second) <= tolerance;

        bool HasInnerRuling(int root) {
            for (int index = 0; index < cells.Count; index++) {
                if (Find(parent, index) != root) continue;
                Box cell = cells[index];
                foreach (Box rule in rules) {
                    cancellationToken.ThrowIfCancellationRequested();
                    // A grid line painted behind an opaque bar is not an inner
                    // cell divider. Unfilled ruled rows do not obscure it.
                    if (cell.Filled && rule.PaintOrder <= cell.PaintOrder) continue;
                    if (acrossRows[root] && Near(rule.Left, rule.Right) && rule.Left > cell.Left + tolerance &&
                        rule.Right < cell.Right - tolerance && rule.Top <= cell.Top + tolerance && rule.Bottom >= cell.Bottom - tolerance ||
                        acrossColumns[root] && Near(rule.Top, rule.Bottom) && rule.Top > cell.Top + tolerance &&
                        rule.Bottom < cell.Bottom - tolerance && rule.Left <= cell.Left + tolerance && rule.Right >= cell.Right - tolerance) return true;
                }
            }
            return false;
        }

        bool HasEdge(Box cell, bool horizontal, double coordinate) {
            foreach (Box rule in rules) {
                cancellationToken.ThrowIfCancellationRequested();
                if (horizontal
                    ? Near(rule.Top, rule.Bottom) && Near(rule.Top, coordinate) &&
                        rule.Left <= cell.Left + tolerance && rule.Right >= cell.Right - tolerance
                    : Near(rule.Left, rule.Right) && Near(rule.Left, coordinate) &&
                        rule.Top <= cell.Top + tolerance && rule.Bottom >= cell.Bottom - tolerance) return true;
            }
            return false;
        }
    }
}
