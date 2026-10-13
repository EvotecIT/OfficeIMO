using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Measures a conservative body top offset once; null leaves foreground placement explicitly unqualified.</summary>
    private static double? GetNativeBodyTextInkTopOffset(WordDocument document, NativeDocumentDefaults defaults,
        NativeFontMap fontMap, WordToPdfOptions? options) {
        var main = document._wordprocessingDocument.MainDocumentPart;
        W.Body? body = main?.Document?.Body;
        if (main == null || body == null) return 0D;
        if (!main.HeaderParts.Any(part => part.Header?.Descendants<W.TextBoxContent>().Any() == true) &&
            !main.FooterParts.Any(part => part.Footer?.Descendants<W.TextBoxContent>().Any() == true)) return 0D;
        double topOffset = 0D;
        foreach (W.Paragraph source in body.Descendants<W.Paragraph>()) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            if (source.Ancestors<W.TextBoxContent>().Any()) continue;
            // Anchored body objects can paint outside the flow frame even with
            // automatic spacing. Table contexts also have effective metrics and
            // clipping that this paragraph-only projection does not inventory.
            if (source.Descendants<W.Drawing>().Any() || source.Descendants<W.Picture>().Any() ||
                source.Ancestors<W.TableCell>().Any()) return null;
            var paragraph = new WordParagraph(document, source);
            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(paragraph);
            PdfCore.PdfParagraphStyle paragraphStyle = CreateNativeParagraphStyle(paragraph, defaults, fontMap);
            if (WordDocumentTraversal.GetListInfo(paragraph).HasValue ||
                HasNativeParagraphBorder(GetNativeEffectiveParagraphBorders(paragraph)) ||
                CreateNativeParagraphPanelStyle(paragraph, paragraphStyle) != null)
                return null;
            NativeLineSpacing spacing = ResolveNativeParagraphLineSpacing(paragraph, styleDefaults, defaults);
            if (spacing.Rule != W.LineSpacingRuleValues.Exact || !spacing.Value.HasValue) continue;
            // These body objects or their decorations have no unclipped ink
            // inventory here. Keep foreground admission explicitly unqualified.
            if (source.Descendants<W.EmbeddedObject>().Any() ||
                source.Descendants<W.FootnoteReference>().Any() || source.Descendants<W.EndnoteReference>().Any() ||
                source.Descendants<W.SimpleField>().Any() || source.Descendants<W.FieldChar>().Any() ||
                WordEquation.GetOccurrences(document, source).Count > 0) return null;
            var canvasStyle = new PdfCore.PdfCanvasTextBoxStyle {
                Background = null, BorderColor = null, BorderWidth = 0D, PaddingX = 0D, PaddingY = 0D,
                FontSize = paragraphStyle.FontSize, LineHeight = paragraphStyle.LineHeight, LineSpacing = paragraphStyle.LineSpacing
            };
            // A page fragment can start with any run. Measure each at the
            // canonical exact-line baseline rather than assuming a font ratio.
            foreach (PdfCore.PdfTextRun run in CreateNativeCellParagraphRuns(paragraph, null,
                    NativeTableStyleDefaults.Empty, defaults, fontMap)) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                if (run.Text.Length == 0) continue;
                if (run.BackgroundColor.HasValue || run.Underline || run.Strike) return null;
                // Layout separators and undecorated spaces have no upward paint.
                if (string.IsNullOrWhiteSpace(run.Text)) continue;
                double? width = fontMap.MeasureText(run);
                if (!width.HasValue) return null;
                if (width.Value <= 0D) continue;
                PdfCore.PdfWriter.CanvasTextBoxMeasurement? measured = fontMap.MeasureCanvasTextBox(new[] { run },
                    width.Value, canvasStyle.LineSpacing!.Value, canvasStyle);
                if (measured == null) return null;
                topOffset = Math.Min(topOffset, measured.TextInkTop);
            }
        }
        return topOffset;
    }
}
