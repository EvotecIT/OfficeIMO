using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfStaticFormRecognizer {
    private static Label? FindLabel(IReadOnlyList<Label> labels, VisualRect field,
        PdfStaticFormEvidenceKind evidence, PdfReadingDirection direction) {
        Label? best = null;
        double bestDistance = double.MaxValue;
        bool ambiguous = false;
        foreach (Label label in labels) {
            double distance = LabelDistance(label, field, evidence, direction);
            if (distance < bestDistance - 0.000001D) {
                best = label;
                bestDistance = distance;
                ambiguous = false;
            } else if (best is not null && Math.Abs(distance - bestDistance) <= 0.000001D) ambiguous = true;
        }
        return ambiguous ? null : best;
    }

    private static double LabelDistance(Label label, VisualRect field,
        PdfStaticFormEvidenceKind evidence, PdfReadingDirection direction) {
        VisualRect bounds = label.Bounds;
        double centerDifference = Math.Abs((bounds.Top + bounds.Bottom) / 2D - (field.Top + field.Bottom) / 2D);
        double distance = double.MaxValue;
        if (centerDifference <= Math.Max(10D, field.Height * 0.65D)) {
            if (bounds.Right <= field.Left && field.Left - bounds.Right <= 120D) distance = field.Left - bounds.Right;
            if ((evidence == PdfStaticFormEvidenceKind.CheckBox || direction == PdfReadingDirection.RightToLeft) &&
                bounds.Left >= field.Right && bounds.Left - field.Right <= 120D) {
                distance = Math.Min(distance, bounds.Left - field.Right);
            }
        }
        if (bounds.Bottom <= field.Top && field.Top - bounds.Bottom <= 30D &&
            bounds.Left <= field.Right && bounds.Right >= field.Left - 15D) {
            distance = Math.Min(distance, 15D + field.Top - bounds.Bottom +
                Math.Abs((bounds.Left + bounds.Right - field.Left - field.Right) / 2D) * 0.1D);
        }
        return distance;
    }

    // Choose after gathering eligible candidates so stream order cannot reserve
    // a label for a farther field. Equal relationships need manual authoring.
    private static List<Candidate> AssignLabels(List<Candidate> candidates,
        Dictionary<int, PdfReadingDirection> directions,
        Action<string, int, string> addDiagnostic, CancellationToken cancellationToken) {
        var accepted = new HashSet<Candidate>();
        foreach (IGrouping<Label, Candidate> group in candidates.GroupBy(static candidate => candidate.Label)) {
            cancellationToken.ThrowIfCancellationRequested();
            Candidate[] choices = group.OrderBy(candidate => LabelDistance(candidate.Label, candidate.Visual,
                candidate.Evidence, directions[candidate.PageNumber])).ToArray();
            double distance = LabelDistance(choices[0].Label, choices[0].Visual,
                choices[0].Evidence, directions[choices[0].PageNumber]);
            bool tied = choices.Length > 1 && Math.Abs(distance - LabelDistance(choices[1].Label,
                choices[1].Visual, choices[1].Evidence, directions[choices[1].PageNumber])) <= 0.000001D;
            if (!tied) accepted.Add(choices[0]);
            if (choices.Length > 1) addDiagnostic("ambiguous-label", choices[0].PageNumber,
                "A label supports at most one field; competing relationships require review.");
        }
        return candidates.Where(accepted.Contains).ToList();
    }

    private static List<VisualRect> GetTableBounds(PdfLogicalPage page, ref long work,
        int maximumWork, CancellationToken cancellationToken) {
        var bounds = new List<VisualRect>(page.Tables.Count);
        foreach (PdfLogicalTable table in page.Tables) {
            ChargeEffectLookup(table.Columns.Count, ref work, maximumWork, cancellationToken);
            if (table.VisualBounds is PdfLogicalVisualBounds visual) {
                bounds.Add(new VisualRect(visual.Left, visual.Top, visual.Right, visual.Bottom));
            } else if (table.Columns.Count > 0) {
                double left = table.Columns.Min(static column => column.From);
                double right = table.Columns.Max(static column => column.To);
                double bottom = Math.Min(table.YTop, table.YBottom);
                double top = Math.Max(table.YTop, table.YBottom);
                if (table.CoordinateSpace == PdfTableCoordinateSpace.VisualTopLeft) {
                    bounds.Add(new VisualRect(left, bottom, right, top));
                } else {
                    PdfVisualBounds mapped = page.TransformBoundsToVisual(left, bottom, right, top);
                    bounds.Add(new VisualRect(mapped.Left, mapped.Top, mapped.Right, mapped.Bottom));
                }
            }
        }
        return bounds;
    }

    private static bool OverlapsDetectedTable(IReadOnlyList<VisualRect> tables, VisualRect bounds) =>
        tables.Any(table => OverlapArea(table, bounds) > 0D);
}
