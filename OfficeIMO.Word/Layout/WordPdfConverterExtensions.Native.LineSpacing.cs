using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        // Retain authored units until the effective paragraph font is resolved. An
        // automatic value is a multiplier, not a height measured with the style font.
        private readonly record struct NativeLineSpacing(double? Value, W.LineSpacingRuleValues? Rule) {
            public double? GridPitch { get; init; }
            public W.VerticalTextAlignmentValues? FontAlignment { get; init; }
            // Word ignores a rule without a value in the same declaration.
            // Inherit the complete pair so its numeric units do not change.
            public NativeLineSpacing Inherit(NativeLineSpacing inherited) =>
                (Value.HasValue ? this : inherited) with { FontAlignment = FontAlignment ?? inherited.FontAlignment };

            public OfficeIMO.Pdf.PdfLineSpacing? ToPdfLineSpacing(double naturalLineHeight) {
                OfficeIMO.Pdf.PdfLineSpacing spacing;
                if (!Value.HasValue) spacing = OfficeIMO.Pdf.PdfLineSpacing.Multiple(naturalLineHeight)
                    .WithFontLineBoxBaseline(naturalLineHeight);
                else if (Rule == null || Rule == W.LineSpacingRuleValues.Auto)
                    spacing = OfficeIMO.Pdf.PdfLineSpacing.Multiple(Math.Max(0.01D, naturalLineHeight * Value.Value / 240D))
                        .WithFontLineBoxBaseline(naturalLineHeight);
                else spacing = Rule == W.LineSpacingRuleValues.AtLeast
                    ? OfficeIMO.Pdf.PdfLineSpacing.AtLeast(Value.Value / 20D, naturalLineHeight).WithFontLineBoxBaseline(naturalLineHeight)
                    // Word's exact-height baseline is four fifths of the line
                    // box in body, column and table exports, regardless of font.
                    : OfficeIMO.Pdf.PdfLineSpacing.Exactly(Value.Value / 20D)
                        .WithFixedLineBoxBaseline(Value.Value / 20D * .8D);
                if (GridPitch.HasValue) spacing = spacing.WithLineGrid(GridPitch.Value);
                // The native PDF layout device uses 600 square pixels per inch.
                // Align each font within its line without rounding line advances.
                if (FontAlignment.HasValue) {
                    var alignment = FontAlignment.Value == W.VerticalTextAlignmentValues.Top ? OfficeIMO.Pdf.PdfFontLineAlignment.Top
                        : FontAlignment.Value == W.VerticalTextAlignmentValues.Center ? OfficeIMO.Pdf.PdfFontLineAlignment.Center
                        : FontAlignment.Value == W.VerticalTextAlignmentValues.Bottom ? OfficeIMO.Pdf.PdfFontLineAlignment.Bottom
                        : OfficeIMO.Pdf.PdfFontLineAlignment.Baseline;
                    spacing = spacing.WithFontAlignment(alignment, 72D / 600D);
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
                if (Rule == null || Rule == W.LineSpacingRuleValues.Auto) {
                    return Math.Max(0.01D, naturalLineHeight * (Value.Value / 240D));
                }
                return fontSize > 0D
                    ? ResolveNativeLineSpacingHeight(Value.Value / 20D, Rule, fontSize, naturalLineHeight)
                    : null;
            }
        }

        private static NativeLineSpacing ReadNativeLineSpacing(W.SpacingBetweenLines? spacing, W.TextAlignment? alignment = null) {
            double? value = double.TryParse(spacing?.Line?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double line) &&
                line >= 0D && !double.IsNaN(line) && !double.IsInfinity(line) ? line : null;
            return new NativeLineSpacing(value, spacing?.LineRule?.Value ?? (value.HasValue ? W.LineSpacingRuleValues.Auto : null)) {
                FontAlignment = alignment?.Val?.Value
            };
        }

        private static NativeLineSpacing ResolveNativeParagraphLineSpacing(
            WordParagraph paragraph, NativeParagraphStyleDefaults styleDefaults,
            NativeDocumentDefaults documentDefaults, NativeTableStyleDefaults table = default) {
            NativeLineSpacing spacing = ReadNativeLineSpacing(paragraph._paragraph.ParagraphProperties?.SpacingBetweenLines,
                paragraph._paragraph.ParagraphProperties?.TextAlignment)
                .Inherit(styleDefaults.LineSpacing.Inherit(table.LineSpacing.Inherit(documentDefaults.LineSpacing)));
            return spacing with { GridPitch = ResolveNativeLineGridPitch(paragraph, styleDefaults, documentDefaults, table) };
        }

        private static double? ResolveNativeLineGridPitch(WordParagraph paragraph, NativeParagraphStyleDefaults styleDefaults,
            NativeDocumentDefaults documentDefaults, NativeTableStyleDefaults table) {
            bool snap = ResolveNativeParagraphSpacing(paragraph, styleDefaults, documentDefaults, table).SnapToGrid ?? true;
            if (!snap) return null;
            if (paragraph._paragraph.Ancestors<W.TableCell>().Any() && ReadNativeOnOff(paragraph._document
                ._wordprocessingDocument.MainDocumentPart?.DocumentSettingsPart?.Settings?
                .GetFirstChild<W.Compatibility>()?.GetFirstChild<W.AdjustLineHeightInTable>()) != true) return null;
            return GetNativeSectionLineGridPitch(paragraph);
        }

        // A paragraph mark or a large run on another line must not impose a
        // paragraph-wide minimum on every visible line. Runs retain their own sizes.
        private static double ResolveNativeParagraphLayoutFontSize(WordParagraph paragraph, NativeDocumentDefaults nativeDefaults,
            NativeParagraphStyleDefaults styleDefaults, NativeTableRunStyleDefaults tableRunStyleDefaults = default) {
            double? minimum = null;
            foreach (WordParagraph run in GetNativeRuns(paragraph)) {
                if (run.IsImage || string.IsNullOrWhiteSpace(run.Text) || IsNativeHiddenTextRun(run, paragraph)) continue;
                double size = ResolveNativeTextRunStyle(run, paragraph, tableRunStyleDefaults, nativeDefaults).FontSize ?? nativeDefaults.FontSize;
                if (size > 0D) minimum = minimum.HasValue ? Math.Min(minimum.Value, size) : size;
            }
            if (minimum.HasValue) return minimum.Value;
            string? markSize = paragraph._paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.FontSize>()?.Val?.Value;
            if (int.TryParse(markSize, NumberStyles.Integer, CultureInfo.InvariantCulture, out int halfPoints) && halfPoints > 0)
                return halfPoints / 2D;
            // With no visible text, hidden run sizes cannot enlarge the mark's line.
            return Math.Max(styleDefaults.FontSize ?? nativeDefaults.FontSize,
                tableRunStyleDefaults.FontSize ?? 0D);
        }
    }
}
