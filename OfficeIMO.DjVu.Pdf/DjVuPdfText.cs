using OfficeIMO.Ocr;
using OfficeIMO.Pdf;

namespace OfficeIMO.DjVu.Pdf;

internal sealed class DjVuPdfTextBudget {
    private readonly DjVuToPdfOptions _options;
    private int _spans;
    internal long Characters { get; private set; }
    internal DjVuPdfTextBudget(DjVuToPdfOptions options) => _options = options;
    internal void Admit(int characters) {
        if (_spans >= _options.MaxTextSpans) throw new DjVuResourceLimitException(nameof(DjVuToPdfOptions.MaxTextSpans));
        if (characters > _options.MaxTextCharacters - Characters) throw new DjVuResourceLimitException(nameof(DjVuToPdfOptions.MaxTextCharacters));
        _spans++; Characters += characters;
    }
}

internal static class DjVuPdfText {
    internal static void AddStored(PdfPageCanvas canvas, DjVuPage page, DjVuTextResult text, DjVuPdfTextBudget budget,
        List<OfficeConversionFidelityDiagnostic> diagnostics, CancellationToken token) {
        var words = Words(text.Zones).Where(w => w.CharacterLength > 0).ToArray();
        bool usable = words.Length != 0;
        int previousEnd = 0;
        foreach (var word in words) {
            token.ThrowIfCancellationRequested();
            if (word.Bounds.Width == 0 || word.Bounds.Height == 0 || word.CharacterOffset < previousEnd) usable = false;
            previousEnd = checked(word.CharacterOffset + word.CharacterLength);
        }
        if (!usable) {
            budget.Admit(text.Text.Length);
            canvas.SearchableText(text.Text, Rectangle(0, 0, page.DisplayWidth * 72.0 / page.Dpi, page.DisplayHeight * 72.0 / page.Dpi));
            DjVuPdfConversionEngine.Add(diagnostics, "djvu.pdf.text-geometry-coarse", "Stored text has no usable word geometry; searchable selection uses the page bounds.", OfficeConversionLossKind.Approximation, page.Number);
            return;
        }
        for (int i = 0; i < words.Length; i++) {
            token.ThrowIfCancellationRequested();
            int start = i == 0 ? 0 : words[i].CharacterOffset;
            int end = i + 1 < words.Length ? words[i + 1].CharacterOffset : text.Text.Length;
            budget.Admit(end - start);
            // Retain original separators in logical order while anchoring each word
            // to its real source geometry. No new OCR provenance is assigned.
            canvas.SearchableText(text.Text.Substring(start, end - start), StoredQuad(page, words[i].Bounds), i);
        }
    }

    internal static bool AddOcr(PdfPageCanvas canvas, DjVuPage page, OcrResult result, int imageWidth, int imageHeight,
        DjVuPdfTextBudget budget, List<OfficeConversionFidelityDiagnostic> diagnostics, CancellationToken token) {
        var spans = (result.Spans ?? Array.Empty<OcrTextSpan>()).Where(s => s != null).OrderBy(s => s.Sequence).ToArray();
        var selected = spans.Where(s => s.Level == OcrTextSpanLevel.Word).ToArray();
        if (selected.Length == 0) selected = spans.Where(s => s.Level == OcrTextSpanLevel.Line).ToArray();
        var request = new OcrRequest { PageNumber = page.Number, PixelWidth = imageWidth, PixelHeight = imageHeight,
            Region = new OcrRegion { Width = page.DisplayWidth * 72.0 / page.Dpi, Height = page.DisplayHeight * 72.0 / page.Dpi },
            RegionCoordinateUnit = OcrCoordinateUnit.Points };
        int placed = 0;
        foreach (var span in selected) {
            token.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(span.Text)) continue;
            if (span.PageNumber.HasValue && span.PageNumber != 1 && span.PageNumber != page.Number || span.Region == null ||
                !OfficeIMO.Pdf.Ocr.PdfOcr.TryConvertRegion(span.Region, span.CoordinateUnit, request, out double x, out double y, out double width, out double height)) {
                DjVuPdfConversionEngine.Add(diagnostics, "djvu.pdf.ocr-geometry-invalid", "An OCR span has invalid page geometry and was not embedded.", OfficeConversionLossKind.Omission, page.Number);
                continue;
            }
            string value = span.Text + (span.Level == OcrTextSpanLevel.Line ? "\n" : " ");
            budget.Admit(value.Length);
            canvas.SearchableText(value, Rectangle(x, y, width, height), placed++);
        }
        if (placed == 0 && !string.IsNullOrWhiteSpace(result.Text))
            DjVuPdfConversionEngine.Add(diagnostics, "djvu.pdf.ocr-geometry-missing", "OCR returned text without usable word or line geometry; the page remains an image without a new text layer.", OfficeConversionLossKind.Omission, page.Number);
        return placed != 0;
    }

    private static IEnumerable<DjVuTextZone> Words(IEnumerable<DjVuTextZone> zones) {
        foreach (var zone in zones) {
            if (zone.Kind == DjVuTextZoneKind.Word) yield return zone;
            else foreach (var word in Words(zone.Children)) yield return word;
        }
    }

    private static PdfSelectionQuad StoredQuad(DjVuPage page, DjVuRectangle bounds) => new(
        Point(page, bounds.X, (long)bounds.Y + bounds.Height),
        Point(page, (long)bounds.X + bounds.Width, (long)bounds.Y + bounds.Height),
        Point(page, (long)bounds.X + bounds.Width, bounds.Y),
        Point(page, bounds.X, bounds.Y));

    private static PdfSelectionPoint Point(DjVuPage page, long x, long y) {
        double dx, dy;
        switch (page.Rotation) {
            case 90: dx = y; dy = x; break;
            case 180: dx = page.Width - x; dy = y; break;
            case 270: dx = page.Height - y; dy = page.Width - x; break;
            default: dx = x; dy = page.Height - y; break;
        }
        return new PdfSelectionPoint(dx * 72.0 / page.Dpi, dy * 72.0 / page.Dpi);
    }

    private static PdfSelectionQuad Rectangle(double x, double y, double width, double height) => new(
        new PdfSelectionPoint(x, y), new PdfSelectionPoint(x + width, y),
        new PdfSelectionPoint(x + width, y + height), new PdfSelectionPoint(x, y + height));
}
