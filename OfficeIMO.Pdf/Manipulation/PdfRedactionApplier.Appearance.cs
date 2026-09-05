namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private static PdfRedactionArea[] BuildPrivacyAppearanceAreas(
        PdfRedactionArea[] areas,
        PdfPageGeometry geometry,
        PdfRedactionApplyOptions options) {
        if (areas.Length == 0) return Array.Empty<PdfRedactionArea>();

        PdfPageBox? pageBox = geometry.EffectiveBox;
        var result = new List<PdfRedactionArea>(areas.Length);
        var merge = new List<PdfRedactionArea>();
        for (int index = 0; index < areas.Length; index++) {
            PdfRedactionArea area = areas[index];
            switch (area.AppearanceMode) {
                case PdfRedactionAppearanceMode.Exact:
                    result.Add(area);
                    break;
                case PdfRedactionAppearanceMode.MergeNearby:
                    merge.Add(ToAppearanceRectangle(area));
                    break;
                case PdfRedactionAppearanceMode.QuantizedWidth:
                    result.Add(QuantizeAppearance(area, pageBox, options.AppearanceWidthQuantum));
                    break;
                case PdfRedactionAppearanceMode.FullLine:
                    if (pageBox is null) throw new InvalidOperationException("Full-line redaction appearance requires a readable effective page box.");
                    result.Add(new PdfRedactionArea(area.PageNumber, pageBox.Left, area.Y, pageBox.Width, area.Height, area.Label, area.ContentScope, area.AppearanceMode));
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(areas), "Redaction area contains an undefined appearance mode.");
            }
        }

        result.AddRange(MergeNearbyAppearanceAreas(merge, options.AppearanceMergeDistance));
        return result.OrderBy(static area => area.Y).ThenBy(static area => area.X).ToArray();
    }

    private static PdfRedactionArea QuantizeAppearance(PdfRedactionArea area, PdfPageBox? pageBox, double quantum) {
        double targetWidth = Math.Ceiling(area.Width / quantum) * quantum;
        double left = area.X - (targetWidth - area.Width) / 2D;
        double right = left + targetWidth;
        if (pageBox is not null) {
            left = Math.Max(pageBox.Left, left);
            right = Math.Min(pageBox.Right, right);
            if (right <= left) throw new InvalidOperationException("Quantized redaction appearance falls outside the effective page box.");
        }
        return new PdfRedactionArea(area.PageNumber, left, area.Y, right - left, area.Height, area.Label, area.ContentScope, area.AppearanceMode);
    }

    private static List<PdfRedactionArea> MergeNearbyAppearanceAreas(List<PdfRedactionArea> areas, double maximumGap) {
        if (areas.Count < 2) return areas;
        var pending = areas.OrderBy(static area => area.Y).ThenBy(static area => area.X).ToList();
        var merged = new List<PdfRedactionArea>(pending.Count);
        while (pending.Count > 0) {
            PdfRedactionArea current = pending[0];
            pending.RemoveAt(0);
            bool changed;
            do {
                changed = false;
                for (int index = 0; index < pending.Count; index++) {
                    PdfRedactionArea candidate = pending[index];
                    if (candidate.PageNumber != current.PageNumber || !SharesVisualLine(current, candidate) || HorizontalGap(current, candidate) > maximumGap) continue;
                    double left = Math.Min(current.X, candidate.X);
                    double bottom = Math.Min(current.Y, candidate.Y);
                    double right = Math.Max(current.Right, candidate.Right);
                    double top = Math.Max(current.Top, candidate.Top);
                    current = new PdfRedactionArea(current.PageNumber, left, bottom, right - left, top - bottom, current.Label, current.ContentScope, current.AppearanceMode);
                    pending.RemoveAt(index);
                    changed = true;
                    break;
                }
            } while (changed);
            merged.Add(current);
        }
        return merged;
    }

    private static bool SharesVisualLine(PdfRedactionArea first, PdfRedactionArea second) {
        double overlap = Math.Min(first.Top, second.Top) - Math.Max(first.Y, second.Y);
        return overlap > 0D && overlap >= Math.Min(first.Height, second.Height) * 0.5D;
    }

    private static double HorizontalGap(PdfRedactionArea first, PdfRedactionArea second) {
        if (first.Right < second.X) return second.X - first.Right;
        if (second.Right < first.X) return first.X - second.Right;
        return 0D;
    }

    private static PdfRedactionArea ToAppearanceRectangle(PdfRedactionArea area) =>
        new(area.PageNumber, area.X, area.Y, area.Width, area.Height, area.Label, area.ContentScope, area.AppearanceMode);
}
