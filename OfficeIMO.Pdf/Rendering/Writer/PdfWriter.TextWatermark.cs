using System.Text;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Synthetic weight paints fill and stroke as one foreground before applying
    // translucent watermark opacity, so overlapping edges retain the same alpha.
    private static string? CreateTextWatermarkGroup(IList<byte[]> objects, PdfCreateOptions options,
        PdfTextWatermark? watermark, string? fontAlias,
        IReadOnlyDictionary<PdfStandardFont, string> pageFonts,
        List<(string Name, int Id)> fonts, List<(string Name, int Id)> xObjects,
        PdfPrintColorTransform? colorTransform, CancellationToken cancellationToken) {
        if (watermark == null || watermark.Opacity <= 0D || watermark.Opacity >= 1D ||
            string.IsNullOrEmpty(fontAlias)) return null;
        PdfStandardFont baseFont = ChooseNormal(watermark.Font);
        if (!BuildTextWatermarkRuns(watermark, options).Any(run =>
            options.NeedsSyntheticBold(ResolvePageTextRunFont(run, baseFont)))) return null;
        var content = new StringBuilder();
        AppendTextWatermark(content, options, watermark, fontAlias!, pageFonts, graphicsStateName: null);
        string foreground = content.ToString();
        if (colorTransform != null) {
            foreground = colorTransform.NormalizeGeneratedContent(foreground, cancellationToken);
        }
        byte[] bytes = PdfEncoding.Latin1GetBytes(foreground);
        var empty = Array.Empty<(string Name, int Id)>();
        string dictionary = PdfTransparencyGroupDictionaryBuilder.BuildStreamDictionary(
            0D, 0D, options.PageWidth, options.PageHeight, bytes.Length,
            FilterPdfResources(foreground, fonts), empty, empty, empty);
        int objectId = AddStreamObject(objects, dictionary, bytes);
        const string resourceName = "/WmText";
        xObjects.Add((resourceName, objectId));
        return resourceName;
    }

    private static void EnsureTextWatermarkFontResources(PdfTextWatermark watermark, PdfCreateOptions options, Func<PdfStandardFont, string, string> ensureFontResource) {
        PdfStandardFont baseFont = ChooseNormal(watermark.Font);
        PdfStandardFont normalFont = ChooseNormal(options.DefaultFont);
        System.Collections.Generic.IReadOnlyList<PdfTextRun> runs = BuildTextWatermarkRuns(watermark, options);
        foreach (PdfTextRun run in runs) {
            PdfStandardFont runFont = ResolvePageTextRunFont(run, baseFont);
            ensureFontResource(runFont, GetStandardFontResourceName(runFont, normalFont));
        }
    }

    private static void AppendTextWatermark(StringBuilder sb, PdfCreateOptions options, PdfTextWatermark watermark, string fontAlias, System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, string> fontResources, string? graphicsStateName) {
        PdfStandardFont baseFont = ChooseNormal(watermark.Font);
        System.Collections.Generic.IReadOnlyList<PdfTextRun> runs = BuildTextWatermarkRuns(watermark, options);
        double textWidth = MeasureTextWatermarkRuns(runs, baseFont, watermark.FontSize, options);
        double angle = watermark.RotationAngle * System.Math.PI / 180D;
        double cos = System.Math.Cos(angle);
        double sin = System.Math.Sin(angle);
        double centerX = options.PageWidth / 2D;
        double centerY = options.PageHeight / 2D;
        double originX = centerX - cos * textWidth / 2D + sin * watermark.FontSize / 2D;
        double originY = centerY - sin * textWidth / 2D - cos * watermark.FontSize / 2D;

        var content = new ContentStreamBuilder(sb)
            .SaveState();
        if (!string.IsNullOrEmpty(graphicsStateName)) {
            content.GraphicsState(graphicsStateName!);
        }

        content
            .BeginText()
            .Font(fontAlias, watermark.FontSize, options.NeedsSyntheticOblique(baseFont), syntheticBold: options.NeedsSyntheticBold(baseFont))
            .FillColor(watermark.Color)
            .TextMatrix(cos, sin, -sin, cos, originX, originY);
        // Keep logical replacement inside the text object so redaction of the
        // watermark removes the visible glyphs and their extraction text together.
        sb.Append("/Span << /ActualText ").Append(PdfSyntaxEscaper.TextString(watermark.Text)).Append(" >> BDC\n");
        foreach (PdfTextRun run in runs) {
            string text = run.Text ?? string.Empty;
            if (text.Length == 0) {
                continue;
            }

            PdfStandardFont runFont = ResolvePageTextRunFont(run, baseFont);
            string runFontResource = ResolvePageTextFontResource(fontResources, runFont);
            double runFontSize = run.FontSize ?? watermark.FontSize;
            content
                .Font(runFontResource, runFontSize, options.NeedsSyntheticOblique(runFont), syntheticBold: options.NeedsSyntheticBold(runFont))
                .ShowText(EncodeTextShowCommand(text, runFont, options), runFontSize, suppressActualText: true);
        }

        sb.Append("EMC\n");
        content.EndText();
        content.RestoreState();
    }

    private static System.Collections.Generic.IReadOnlyList<PdfTextRun> BuildTextWatermarkRuns(PdfTextWatermark watermark, PdfCreateOptions options) {
        PdfStandardFont baseFont = ChooseNormal(watermark.Font);
        var run = new PdfTextRun(
            watermark.Text,
            bold: watermark.Bold,
            underline: false,
            color: watermark.Color,
            italic: watermark.Italic,
            strike: false,
            fontSize: watermark.FontSize,
            font: baseFont);
        return NormalizeFallbackRuns(new[] { run }, baseFont, options);
    }

    private static double MeasureTextWatermarkRuns(System.Collections.Generic.IReadOnlyList<PdfTextRun> runs, PdfStandardFont baseFont, double fontSize, PdfCreateOptions options) {
        double width = 0D;
        foreach (PdfTextRun run in runs) {
            width += MeasureRichText(run.Text ?? string.Empty, ResolvePageTextRunFont(run, baseFont), run.FontSize ?? fontSize, run.Baseline, options);
        }

        return width;
    }

    private static PdfStandardFont GetTextWatermarkFont(PdfTextWatermark watermark) {
        PdfStandardFont normal = ChooseNormal(watermark.Font);
        if (watermark.Bold && watermark.Italic) {
            return ChooseBoldItalic(normal);
        }

        if (watermark.Bold) {
            return ChooseBold(normal);
        }

        if (watermark.Italic) {
            return ChooseItalic(normal);
        }

        return normal;
    }

}
