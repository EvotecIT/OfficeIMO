using System.Globalization;

namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private const double NativeLayoutDevicePointsPerPixel = 72D / 600D;
        // Retain authored units until the effective paragraph font is resolved. An
        // automatic value is a multiplier, not a height measured with the style font.
        private readonly record struct NativeLineSpacing(double? Value, string? Rule) {
            public double? GridPitch { get; init; }
            public string? FontAlignment { get; init; }
            // WordPerfect-compatible line height: the natural box is the font size plus a
            // fixed leading rather than each font's own line metrics.
            public bool FixedFontLineBox { get; init; }
            // Word ignores a rule without a value in the same declaration.
            // Inherit the complete pair so its numeric units do not change.
            public NativeLineSpacing Inherit(NativeLineSpacing inherited) =>
                (Value.HasValue ? this : inherited) with { FontAlignment = FontAlignment ?? inherited.FontAlignment };

            public OfficeIMO.Pdf.PdfLineSpacing? ToPdfLineSpacing(double naturalLineHeight) {
                OfficeIMO.Pdf.PdfLineSpacing spacing;
                if (!Value.HasValue) spacing = OfficeIMO.Pdf.PdfLineSpacing.Multiple(naturalLineHeight)
                    .WithFontLineBoxBaseline(naturalLineHeight);
                else if (Rule == null || Rule == "auto")
                    spacing = OfficeIMO.Pdf.PdfLineSpacing.Multiple(Math.Max(0.01D, naturalLineHeight * Value.Value / 240D))
                        .WithFontLineBoxBaseline(naturalLineHeight);
                else spacing = Rule == "atLeast"
                    ? OfficeIMO.Pdf.PdfLineSpacing.AtLeast(global::OfficeIMO.OfficeLength.Points(Value.Value / 20D), naturalLineHeight).WithFontLineBoxBaseline(naturalLineHeight)
                    // Word's exact-height baseline is four fifths of the line
                    // box in body, column and table exports, regardless of font.
                    : OfficeIMO.Pdf.PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(Value.Value / 20D))
                        .WithFixedLineBoxBaseline(Value.Value / 20D * .8D);
                if (GridPitch.HasValue) spacing = spacing.WithLineGrid(GridPitch.Value);
                if (FixedFontLineBox && !spacing.IsExact) spacing = spacing.WithFixedFontLeading(NativeWordPerfectLeading);
                // The native PDF layout device uses 600 square pixels per inch.
                // Align each font within its line without rounding line advances.
                if (FontAlignment != null) {
                    var alignment = FontAlignment == "top" ? OfficeIMO.Pdf.PdfFontLineAlignment.Top
                        : FontAlignment == "center" ? OfficeIMO.Pdf.PdfFontLineAlignment.Center
                        : FontAlignment == "bottom" ? OfficeIMO.Pdf.PdfFontLineAlignment.Bottom
                        : OfficeIMO.Pdf.PdfFontLineAlignment.Baseline;
                    spacing = spacing.WithFontAlignment(alignment, NativeLayoutDevicePointsPerPixel);
                }
                return spacing;
            }

            public double? Resolve(double fontSize, double naturalLineHeight) {
                if (GridPitch.HasValue && fontSize > 0D)
                    return ToPdfLineSpacing(naturalLineHeight)!.GetAdvance(fontSize) / fontSize;
                // An undeclared line value still uses the effective font's
                // natural box, just as ToPdfLineSpacing does for visible text.
                // Empty paragraphs and object spacing measure this value
                // directly instead of entering the rich-text layout engine.
                if (!Value.HasValue) return naturalLineHeight;
                if (Rule == null || Rule == "auto") {
                    return Math.Max(0.01D, naturalLineHeight * (Value.Value / 240D));
                }
                return fontSize > 0D
                    ? ResolveNativeLineSpacingHeight(Value.Value / 20D, Rule, fontSize, naturalLineHeight)
                    : null;
            }
        }

        private static NativeLineSpacing ReadNativeLineSpacing(OfficeXmlNode spacing, OfficeXmlNode alignment = default) {
            double? value = OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(spacing.WAttr("line"), NumberStyles.Float, CultureInfo.InvariantCulture, out double line) &&
                line >= 0D && !double.IsNaN(line) && !double.IsInfinity(line) ? line : null;
            return new NativeLineSpacing(value, spacing.WAttr("lineRule") ?? (value.HasValue ? "auto" : null)) {
                FontAlignment = alignment.WVal()
            };
        }

        private static NativeLineSpacing ResolveNativeParagraphLineSpacing(
            WordParagraph paragraph, NativeParagraphStyleDefaults styleDefaults,
            NativeDocumentDefaults documentDefaults, NativeTableStyleDefaults? table = null) {
            table ??= NativeTableStyleDefaults.Empty;
            NativeLineSpacing spacing = ReadNativeLineSpacing(paragraph.Node().WChild("pPr").WChild("spacing"),
                paragraph.Node().WChild("pPr").WChild("textAlignment"))
                .Inherit(styleDefaults.LineSpacing.Inherit(table.LineSpacing.Inherit(documentDefaults.LineSpacing)));
            return spacing with {
                GridPitch = ResolveNativeLineGridPitch(paragraph, styleDefaults, documentDefaults, table),
                FixedFontLineBox = UsesNativeWordPerfectLineHeight(paragraph._document)
            };
        }

        private static double? ResolveNativeLineGridPitch(WordParagraph paragraph, NativeParagraphStyleDefaults styleDefaults,
            NativeDocumentDefaults documentDefaults, NativeTableStyleDefaults table) {
            bool snap = ResolveNativeParagraphSpacing(paragraph, styleDefaults, documentDefaults, table).SnapToGrid ?? true;
            if (!snap) return null;
            // A projected frame copy keeps the line grid of the section its paragraph is in.
            if (IsNativeFrameCopy(paragraph)) return GetNativeSectionLineGridPitch(ResolveNativeFrameSourceParagraph(paragraph), documentDefaults.NoteContentSection);
            if (paragraph.Node().Ancestor(WordXml.W, "tc").Exists && ReadNativeOnOff(paragraph._document
                .SettingsNode().WChild("compat").WChild("adjustLineHeightInTable")) != true) return null;
            return GetNativeSectionLineGridPitch(paragraph, documentDefaults.NoteContentSection);
        }

        // A note's special rule character has no visible text, but retains its
        // source font metrics. It is not an empty paragraph or a spacing run.
        private static bool IsNativeParagraphMetricBearingRun(WordParagraph run, WordParagraph paragraph, bool noteContent) =>
            !run.IsImage && !IsNativeHiddenTextRun(run, paragraph) &&
            (!string.IsNullOrWhiteSpace(run.Text) || noteContent &&
                run.EffectiveRunContentNodes().Any(child => child.IsW("separator") || child.IsW("continuationSeparator")));

        // A paragraph mark or a large run on another line must not impose a
        // paragraph-wide minimum on every content line. Runs retain their own sizes.
        private static double ResolveNativeParagraphLayoutFontSize(WordParagraph paragraph, NativeDocumentDefaults nativeDefaults,
            NativeParagraphStyleDefaults styleDefaults, NativeTableRunStyleDefaults tableRunStyleDefaults = default) {
            double? minimum = null;
            foreach (WordParagraph run in GetNativeRuns(paragraph)) {
                if (!IsNativeParagraphMetricBearingRun(run, paragraph, nativeDefaults.NoteContentSection != null)) continue;
                double size = ResolveNativeTextRunStyle(run, paragraph, tableRunStyleDefaults, nativeDefaults).FontSize ?? nativeDefaults.FontSize;
                if (size > 0D) minimum = minimum.HasValue ? Math.Min(minimum.Value, size) : size;
            }
            if (minimum.HasValue) return minimum.Value;
            string? markSize = paragraph.Node().WChild("pPr").WChild("rPr").WChild("sz").WVal();
            if (int.TryParse(markSize, NumberStyles.Integer, CultureInfo.InvariantCulture, out int halfPoints) && halfPoints > 0)
                return halfPoints / 2D;
            // With no visible text, hidden run sizes cannot enlarge the mark's line.
            return Math.Max(styleDefaults.FontSize ?? nativeDefaults.FontSize,
                tableRunStyleDefaults.FontSize ?? 0D);
        }
    }
}
