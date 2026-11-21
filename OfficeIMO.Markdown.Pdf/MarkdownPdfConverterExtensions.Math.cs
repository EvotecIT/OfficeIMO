using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.PdfTextRun;

namespace OfficeIMO.Markdown.Pdf;

// TeX math from MathInline, MathBlock and math semantic fences, laid out by the shared OfficeIMO math engine.
public static partial class MarkdownPdfConverterExtensions {
    private const int MathFontByteLimit = 32 * 1024 * 1024;
    private const double InlineMathRasterScale = 4D;

    internal sealed class MathLayout {
        private readonly OfficeFontFaceCollection _fonts;

        private MathLayout(OfficeFontFaceCollection fonts, double fontSize) {
            _fonts = fonts;
            FontSize = fontSize;
        }

        internal double FontSize { get; }

        // Loads the installed math font once per export and lets PDF drawings embed the same face that measured the layout.
        internal static MathLayout? Prepare(MarkdownDocument document, PdfCore.PdfCreateOptions pdfOptions, MarkdownPdfExportOptions options) {
            var finder = new MathFinder();
            finder.Visit(document);
            if (!finder.Found) return null;

            string? reason = null;
            var fonts = new OfficeFontFaceCollection();
            if (!options.ResourcePolicy.AllowSystemFontEmbedding) {
                reason = "The resource policy does not allow installed fonts, and a mathematical font is required to lay out TeX math.";
            } else if (!fonts.TryAddInstalledFamily("math", OfficeFontFaceDescriptor.Regular, MathFontByteLimit,
                           options.CancellationToken, out _, out string? error, MathFontByteLimit)) {
                reason = "No installed mathematical font (STIX Two Math, Cambria Math, Latin Modern Math or DejaVu Math TeX Gyre) could be loaded" +
                    (string.IsNullOrWhiteSpace(error) ? "." : ": " + error);
            } else {
                try {
                    pdfOptions.UseRenderingProfile(new OfficeRenderingProfile("officeimo-markdown-math", fonts), OfficeRenderingProfileApplyMode.Overlay);
                } catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or NotSupportedException) {
                    reason = "The mathematical font could not be registered with the PDF options: " + ex.Message;
                }
            }

            if (reason != null) {
                AddWarning(options, "math-font-unavailable", "math", reason + " Math is shown as TeX source.");
                return null;
            }

            return new MathLayout(fonts, pdfOptions.DefaultFontSize.ToPoints());
        }

        internal bool TryRender(string tex, bool display, double fontSize, OfficeColor color, MarkdownPdfExportOptions options,
            out OfficeDrawing drawing, out OfficeMathLayoutMetrics metrics) {
            drawing = null!;
            metrics = default;
            OfficeMathExpression expression;
            try {
                expression = OfficeMathMarkup.FromLatex(tex);
            } catch (Exception ex) when (ex is OfficeMathParseException or FormatException or ArgumentException) {
                AddWarning(options, "math-unsupported", tex, "The TeX math could not be read by the OfficeIMO math engine and is shown as source. " + ex.Message);
                return false;
            }

            var renderOptions = new OfficeMathRenderOptions {
                Font = new OfficeFontInfo("math", fontSize),
                Color = color,
                Padding = 0D,
                RuleGap = Math.Max(0.75D, fontSize / 16D),
                RuleThickness = Math.Max(0.6D, fontSize / 20D),
                MatrixGap = Math.Max(4D, fontSize * 0.6D),
                DisplayStyle = display,
                Fonts = _fonts,
                TokenPaintText = PaintTexIdentifier
            };
            metrics = OfficeMathRenderer.Measure(expression, renderOptions, options.CancellationToken);
            drawing = OfficeMathRenderer.Render(expression, renderOptions, options.CancellationToken);
            return drawing.Width > 0D && drawing.Height > 0D;
        }

        // TeX sets letters of an implicit product (mc, dx) and lowercase Greek in math italic; operator names stay upright.
        private static string? PaintTexIdentifier(OfficeMathExpression expression) {
            if (expression.Kind != OfficeMathKind.Identifier || string.IsNullOrEmpty(expression.Text)) return null;
            string text = expression.Text!;
            var painted = new StringBuilder(text.Length * 2);
            foreach (char value in text) {
                bool italic = value is >= 'a' and <= 'z' or >= 'A' and <= 'Z' or >= 'α' and <= 'ω';
                painted.Append(italic ? OfficeMathTextTransform.MathAuto(value.ToString()) : value.ToString());
            }

            string result = painted.ToString();
            return result == text ? null : result;
        }
    }

    private sealed class MathFinder : MarkdownVisitor {
        internal bool Found { get; private set; }

        protected override void VisitInline(MarkdownInline inline) {
            if (inline is MathInline) Found = true;
            else if (!Found) base.VisitInline(inline);
        }

        protected override void VisitBlock(MarkdownBlock block) {
            if (block is MathBlock) Found = true;
            else if (!Found) base.VisitBlock(block);
        }

        protected override void VisitSemanticFencedBlock(SemanticFencedBlock block) {
            if (IsMathSemanticFence(block)) Found = true;
        }
    }

    // Math fences carry TeX; AsciiMath fences from AsciiDoc stay code panels.
    private static bool IsMathSemanticFence(SemanticFencedBlock semantic) =>
        string.Equals(semantic.SemanticKind, MarkdownSemanticKinds.Math, StringComparison.OrdinalIgnoreCase) &&
        !string.Equals(semantic.Language, "asciimath", StringComparison.OrdinalIgnoreCase);

    private static bool TryRenderDisplayMath(PdfCore.PdfDocument pdf, string tex, string? caption, MarkdownPdfExportOptions options, MarkdownPdfStyle visualTheme) {
        MathLayout? layout = options.Math;
        if (layout == null || !layout.TryRender(tex, display: true, layout.FontSize * 1.1D, OfficeColor.Black, options, out OfficeDrawing drawing, out _)) {
            return false;
        }

        UsePaintedMathGlyphs(drawing);
        MarkdownPdfFigureStyle figureStyle = visualTheme.FigureStyleSnapshot;
        PdfCore.PdfDrawingStyle drawingStyle = figureStyle.DrawingStyleSnapshot;
        drawingStyle.Decorative = false;
        drawingStyle.ConstrainToContentWidth = true;
        drawingStyle.AlternativeText = tex;
        pdf.Drawing(drawing, OfficeAlignment.Center, spacingBefore: layout.FontSize * 0.5D, spacingAfter: layout.FontSize * 0.5D, style: drawingStyle);
        RenderCaption(pdf, caption, figureStyle);
        return true;
    }

    // PDF drawing text paints the logical letters; math italic lives only in the painted glyph text,
    // so the PDF receives the painted glyphs and the formula keeps its TeX source as alternative text.
    private static void UsePaintedMathGlyphs(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingText text && !string.Equals(text.Text, text.RasterText, StringComparison.Ordinal) && text.RasterText.Length > 0) {
                text.SetPaintedText(text.RasterText, text.RasterText);
            } else if (element is OfficeDrawingGroup group) {
                UsePaintedMathGlyphs(group.Drawing);
            }
        }
    }

    private static void RenderMathBlock(PdfCore.PdfDocument pdf, MathBlock math, MarkdownPdfExportOptions options, MarkdownPdfStyle visualTheme) {
        if (TryRenderDisplayMath(pdf, math.Content, caption: null, options, visualTheme)) return;
        RenderSemanticSourcePanel(pdf, "math", math.Content, caption: null, options, visualTheme);
    }

    // Inline math is painted from the same vector layout at four times the reading resolution, sitting on the text baseline.
    private static bool TryCreateInlineMathRun(MathInline math, InlineStyle style, out PdfTextRun run) {
        run = null!;
        MarkdownPdfExportOptions? options = style.Anchors?.Options;
        MathLayout? layout = options?.Math;
        double fontSize = style.FontSize ?? layout?.FontSize ?? 11D;
        if (layout == null || !layout.TryRender(math.Source, math.IsDisplay, fontSize, style.Color ?? OfficeColor.Black, options!,
                out OfficeDrawing drawing, out OfficeMathLayoutMetrics metrics)) {
            return false;
        }

        byte[] png;
        try {
            png = OfficeDrawingRasterRenderer.ToPng(drawing, InlineMathRasterScale);
        } catch (Exception ex) when (ex is NotSupportedException or InvalidOperationException or ArgumentException) {
            AddWarning(options!, "math-unsupported", math.Source, "The inline math could not be painted and is shown as source. " + ex.Message);
            return false;
        }

        double descent = Math.Max(0D, drawing.Height - metrics.Baseline);
        run = PdfTextRun.Inline(new PdfCore.PdfInlineImage(png, OfficeLength.Points(drawing.Width), OfficeLength.Points(drawing.Height),
            math.Source, OfficeImageFit.Contain, OfficeLength.Points(-descent)));
        return true;
    }

    private static PdfTextRun CreateMathSourceRun(MathInline math, InlineStyle style) =>
        CreateRun(math.Source, style.With(italic: true));
}
