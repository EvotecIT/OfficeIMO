using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word.Html;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        private static string? ResolveDefaultFontFamily(MarkdownToWordOptions options) {
            if (options == null) {
                return null;
            }

            return FontResolver.Resolve(options.FontFamily) ?? options.FontFamily;
        }

        private static void ApplyBlockParagraphFormatting(WordParagraph paragraph, int quoteDepth, Omd.ColumnAlignment alignment) {
            if (quoteDepth > 0) {
                paragraph.IndentationBefore = IndentTwipsPerLevel * quoteDepth;
            }

            ApplyAlignment(alignment, paragraph);
        }

        private static Omd.MarkdownVisualTheme? ResolveTheme(MarkdownToWordOptions options) =>
            Omd.MarkdownVisualTheme.ResolveOrDefault(options.ThemeSnapshot, options.ApplyDefaultTheme);

        private static void ApplyHeadingTheme(WordParagraph paragraph, MarkdownToWordOptions options) {
            Omd.MarkdownVisualTheme? theme = ResolveTheme(options);
            if (theme == null) {
                return;
            }

            Omd.MarkdownVisualPalette palette = theme.PaletteSnapshot;
            foreach (var run in paragraph.GetRuns()) {
                if (run.IsHyperLink) {
                    continue;
                }

                run.SetColorHex(palette.Heading.ToRgbHex());
            }
        }

        private static void ApplyCodeTheme(WordParagraph paragraph, MarkdownToWordOptions options) {
            Omd.MarkdownVisualTheme? theme = ResolveTheme(options);
            if (theme == null) {
                return;
            }

            Omd.MarkdownVisualPalette palette = theme.PaletteSnapshot;
            if (palette.CodeBackground.A > 0) {
                paragraph.ShadingFillColorHex = palette.CodeBackground.ToRgbHex();
            }

            paragraph.Borders.LeftStyle = BorderValues.Single;
            paragraph.Borders.LeftColorHex = palette.Border.ToRgbHex();
            paragraph.Borders.LeftSize = 4;
            foreach (var run in paragraph.GetRuns()) {
                run.SetColorHex(palette.Text.ToRgbHex());
            }
        }

        private static void ApplyBodyTextTheme(WordParagraph paragraph, MarkdownToWordOptions options) {
            Omd.MarkdownVisualTheme? theme = ResolveTheme(options);
            if (theme == null) {
                return;
            }

            string textHex = theme.PaletteSnapshot.Text.ToRgbHex();
            foreach (var run in paragraph.GetRuns()) {
                run.SetColorHex(textHex);
            }
        }

        private static void ApplyCalloutTitleTheme(WordParagraph paragraph, MarkdownToWordOptions options) {
            Omd.MarkdownVisualTheme? theme = ResolveTheme(options);
            if (theme == null) {
                return;
            }

            Omd.MarkdownVisualPalette palette = theme.PaletteSnapshot;
            if (palette.Surface.A > 0) {
                paragraph.ShadingFillColorHex = palette.Surface.ToRgbHex();
            }

            paragraph.Borders.LeftStyle = BorderValues.Single;
            paragraph.Borders.LeftColorHex = palette.Accent.ToRgbHex();
            paragraph.Borders.LeftSize = 8;
            foreach (var run in paragraph.GetRuns()) {
                run.SetColorHex(palette.Accent.ToRgbHex());
                run.SetBold();
            }
        }

        private static void ApplyTableTheme(WordTable table, MarkdownToWordOptions options, bool hasHeaderRow) {
            Omd.MarkdownVisualTheme? theme = ResolveTheme(options);
            if (theme == null) {
                return;
            }

            Omd.MarkdownVisualPalette palette = theme.PaletteSnapshot;
            Omd.MarkdownTableVisualStyle tableStyle = theme.TableSnapshot;
            bool hasVisibleBorder = palette.Border.A > 0 && IsPositiveFinite(tableStyle.BorderWidth);
            string? borderHex = hasVisibleBorder ? palette.Border.ToRgbHex() : null;
            DocumentFormat.OpenXml.UInt32Value borderSize = hasVisibleBorder ? ToWordBorderSize(tableStyle.BorderWidth) : 0U;
            BorderValues borderStyle = hasVisibleBorder ? BorderValues.Single : BorderValues.None;
            short horizontalPadding = ToWordCellMarginWidth(tableStyle.CellPaddingX);
            short verticalPadding = ToWordCellMarginWidth(tableStyle.CellPaddingY);

            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                bool header = hasHeaderRow && rowIndex == 0;
                bool stripe = !header && tableStyle.UseRowStripes && rowIndex % 2 == 0;
                foreach (var cell in table.Rows[rowIndex].Cells) {
                    cell.MarginLeftWidth = horizontalPadding;
                    cell.MarginRightWidth = horizontalPadding;
                    cell.MarginTopWidth = verticalPadding;
                    cell.MarginBottomWidth = verticalPadding;
                    cell.Borders.TopStyle = borderStyle;
                    cell.Borders.BottomStyle = borderStyle;
                    cell.Borders.LeftStyle = borderStyle;
                    cell.Borders.RightStyle = borderStyle;
                    cell.Borders.TopSize = borderSize;
                    cell.Borders.BottomSize = borderSize;
                    cell.Borders.LeftSize = borderSize;
                    cell.Borders.RightSize = borderSize;
                    cell.Borders.TopColorHex = borderHex;
                    cell.Borders.BottomColorHex = borderHex;
                    cell.Borders.LeftColorHex = borderHex;
                    cell.Borders.RightColorHex = borderHex;

                    if (header && tableStyle.EmphasizeHeader) {
                        ApplyCellFillColor(cell, palette.TableHeaderBackground);
                        ApplyCellRunColor(cell, palette.TableHeaderText);
                    } else if (stripe) {
                        ApplyCellFillColor(cell, palette.TableStripeBackground);
                    }
                }
            }
        }

        private static DocumentFormat.OpenXml.UInt32Value ToWordBorderSize(double borderWidth) {
            if (!IsPositiveFinite(borderWidth)) {
                return 0U;
            }

            double value = borderWidth;
            uint size = (uint)Math.Max(2, Math.Min(96, Math.Round(value * 8, MidpointRounding.AwayFromZero)));
            return size;
        }

        private static short ToWordCellMarginWidth(double padding) {
            if (!IsPositiveFinite(padding)) {
                return 0;
            }

            return (short)Math.Max(0, Math.Min(short.MaxValue, Math.Round(padding * 20, MidpointRounding.AwayFromZero)));
        }

        private static bool IsPositiveFinite(double value) =>
            !double.IsNaN(value) && !double.IsInfinity(value) && value > 0;

        private static void ApplyCellFillColor(WordTableCell cell, OfficeColor color) {
            if (color.A > 0) {
                cell.ShadingFillColorHex = color.ToRgbHex();
            }
        }

        private static void ApplyCellRunColor(WordTableCell cell, OfficeColor color) {
            if (color.A == 0) {
                return;
            }

            string colorHex = color.ToRgbHex();
            foreach (var paragraph in cell.Paragraphs) {
                foreach (var run in paragraph.GetRuns()) {
                    run.SetColorHex(colorHex);
                }
            }
        }

        private static void ApplyAlignment(Omd.ColumnAlignment align, WordParagraph para) {
            switch (align) {
                case Omd.ColumnAlignment.Left: para.ParagraphAlignment = JustificationValues.Left; break;
                case Omd.ColumnAlignment.Center: para.ParagraphAlignment = JustificationValues.Center; break;
                case Omd.ColumnAlignment.Right: para.ParagraphAlignment = JustificationValues.Right; break;
            }
        }
    }
}
