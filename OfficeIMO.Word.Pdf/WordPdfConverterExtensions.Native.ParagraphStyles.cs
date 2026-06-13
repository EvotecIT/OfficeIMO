using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static PdfCore.PdfParagraphStyle CreateNativeParagraphStyle(WordParagraph paragraph) {
            var style = new PdfCore.PdfParagraphStyle();
            if (paragraph.LineSpacingBeforePoints.HasValue) {
                style.SpacingBefore = paragraph.LineSpacingBeforePoints.Value;
            }

            if (paragraph.LineSpacingAfterPoints.HasValue) {
                style.SpacingAfter = paragraph.LineSpacingAfterPoints.Value;
            }

            if (paragraph.IndentationBeforePoints.HasValue) {
                style.LeftIndent = paragraph.IndentationBeforePoints.Value;
            }

            if (paragraph.IndentationAfterPoints.HasValue) {
                style.RightIndent = paragraph.IndentationAfterPoints.Value;
            }

            if (paragraph.IndentationFirstLinePoints.HasValue) {
                style.FirstLineIndent = paragraph.IndentationFirstLinePoints.Value;
            }

            if (paragraph.IndentationHangingPoints.HasValue) {
                double hangingIndent = paragraph.IndentationHangingPoints.Value;
                if (style.LeftIndent < hangingIndent) {
                    style.LeftIndent = hangingIndent;
                }

                style.FirstLineIndent = -hangingIndent;
            }

            double fontSize = paragraph.FontSize.HasValue && paragraph.FontSize.Value > 0 ? paragraph.FontSize.Value : 11D;
            style.LineHeight = ResolveNativeParagraphLineHeight(paragraph, fontSize);

            if (!paragraph.LineSpacingAfterPoints.HasValue) {
                style.SpacingAfter = NativeDefaultParagraphSpacingAfter;
            }

            double? defaultTabStopWidth = GetNativeDefaultTabStopWidth(paragraph);
            if (defaultTabStopWidth.HasValue) {
                style.DefaultTabStopWidth = defaultTabStopWidth.Value;
            }

            foreach (WordTabStop tabStop in paragraph.TabStops
                .Where(tabStop => tabStop.Position > 0)
                .OrderBy(tabStop => tabStop.Position)) {
                double? position = ConvertNativeTwipsToPoints(tabStop.Position);
                if (!position.HasValue) {
                    continue;
                }

                double framePosition = position.Value - style.LeftIndent;
                if (framePosition <= 0D) {
                    continue;
                }

                style.TabStops.Add(new PdfCore.PdfTabStop(
                    framePosition,
                    MapNativeTabAlignment(tabStop.Alignment),
                    MapNativeTabLeader(tabStop.Leader)));
            }

            style.KeepTogether = paragraph.KeepLinesTogether;
            style.KeepWithNext = paragraph.KeepWithNext;
            style.WidowControl = paragraph.AvoidWidowAndOrphan;
            return style;
        }

        private static double ResolveNativeParagraphLineHeight(WordParagraph paragraph, double fontSize) {
            if (paragraph.LineSpacing.HasValue && paragraph.LineSpacingRule == W.LineSpacingRuleValues.Auto) {
                return Math.Max(0.01D, paragraph.LineSpacing.Value / 240D);
            }

            if (paragraph.LineSpacingPoints.HasValue && fontSize > 0D) {
                return paragraph.LineSpacingPoints.Value / fontSize;
            }

            return NativeDefaultParagraphLineHeight;
        }

        private static double? GetNativeDefaultTabStopWidth(WordParagraph paragraph) {
            int firstTabStop = paragraph.TabStops
                .Where(tabStop => tabStop.Position > 0)
                .Select(tabStop => tabStop.Position)
                .DefaultIfEmpty(0)
                .Min();

            return firstTabStop > 0 ? ConvertNativeTwipsToPoints(firstTabStop) : null;
        }

        private static PdfCore.PdfTabLeaderStyle MapNativeTabLeader(W.TabStopLeaderCharValues leader) {
            if (leader == W.TabStopLeaderCharValues.Dot || leader == W.TabStopLeaderCharValues.MiddleDot || leader == W.TabStopLeaderCharValues.Heavy) {
                return PdfCore.PdfTabLeaderStyle.Dots;
            }

            if (leader == W.TabStopLeaderCharValues.Hyphen) {
                return PdfCore.PdfTabLeaderStyle.Hyphens;
            }

            if (leader == W.TabStopLeaderCharValues.Underscore) {
                return PdfCore.PdfTabLeaderStyle.Underscores;
            }

            return PdfCore.PdfTabLeaderStyle.None;
        }

        private static PdfCore.PdfTabAlignment MapNativeTabAlignment(W.TabStopValues alignment) {
            if (alignment == W.TabStopValues.Center) {
                return PdfCore.PdfTabAlignment.Center;
            }

            if (alignment == W.TabStopValues.Right) {
                return PdfCore.PdfTabAlignment.Right;
            }

            if (alignment == W.TabStopValues.Decimal) {
                return PdfCore.PdfTabAlignment.DecimalSeparator;
            }

            return PdfCore.PdfTabAlignment.Left;
        }

        private static PdfCore.PanelStyle? CreateNativeParagraphPanelStyle(WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle) {
            PdfCore.PdfColor? background = ParseNativeColor(paragraph.ShadingFillColorHex);
            (PdfCore.PdfColor? Color, double Width)? border = GetNativeUniformParagraphBorder(paragraph.Borders);
            bool renderAsRule = !background.HasValue &&
                (HasNativeOnlyTopParagraphBorder(paragraph.Borders) || HasNativeOnlyBottomParagraphBorder(paragraph.Borders));
            if (renderAsRule) {
                return null;
            }

            bool hasParagraphBorder = HasNativeParagraphBorder(paragraph.Borders);
            if (!background.HasValue && border == null && !hasParagraphBorder) {
                return null;
            }

            var style = new PdfCore.PanelStyle {
                Background = background,
                BorderColor = border?.Color,
                BorderWidth = border?.Width ?? 0D,
                PaddingX = 6,
                PaddingY = 4,
                SpacingBefore = paragraphStyle.SpacingBefore,
                SpacingAfter = paragraphStyle.SpacingAfter ?? 6D,
                Align = MapNativeParagraphAlign(paragraph.ParagraphAlignment, allowJustify: false)
            };

            if (border == null && hasParagraphBorder) {
                style.TopBorder = CreateNativePanelBorder(paragraph.Borders.TopStyle, paragraph.Borders.TopColorHex, paragraph.Borders.TopSize);
                style.RightBorder = CreateNativePanelBorder(paragraph.Borders.RightStyle, paragraph.Borders.RightColorHex, paragraph.Borders.RightSize);
                style.BottomBorder = CreateNativePanelBorder(paragraph.Borders.BottomStyle, paragraph.Borders.BottomColorHex, paragraph.Borders.BottomSize);
                style.LeftBorder = CreateNativePanelBorder(paragraph.Borders.LeftStyle, paragraph.Borders.LeftColorHex, paragraph.Borders.LeftSize);
            }

            return style;
        }

        private static bool IsNativeHorizontalRuleParagraph(WordParagraph paragraph, IReadOnlyList<WordParagraph> runs, string content) {
            if (!string.IsNullOrEmpty(content) ||
                paragraph.Image != null ||
                paragraph.Shape != null ||
                paragraph.TextBox != null ||
                runs.Any(run => run.IsImage || !string.IsNullOrEmpty(run.Text))) {
                return false;
            }

            return HasNativeOnlyBottomParagraphBorder(paragraph.Borders);
        }

        private static PdfCore.PdfHorizontalRuleStyle? CreateNativeHorizontalRuleStyle(WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle) {
            WordParagraphBorders borders = paragraph.Borders;
            if (!HasNativeBorder(borders.BottomStyle)) {
                return null;
            }

            return new PdfCore.PdfHorizontalRuleStyle {
                Thickness = (borders.BottomSize?.Value ?? 4U) / 8D,
                Color = ParseNativeColor(NormalizeNativeBorderColor(borders.BottomColorHex)) ?? PdfCore.PdfColor.Black,
                SpacingBefore = paragraphStyle.SpacingBefore,
                SpacingAfter = paragraphStyle.SpacingAfter ?? (borders.BottomSpace?.Value ?? 6U),
                KeepWithNext = paragraphStyle.KeepWithNext
            };
        }

        private static PdfCore.PdfHorizontalRuleStyle? CreateNativeBottomBorderRuleStyle(WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle) {
            WordParagraphBorders borders = paragraph.Borders;
            if (!HasNativeOnlyBottomParagraphBorder(borders)) {
                return null;
            }

            return new PdfCore.PdfHorizontalRuleStyle {
                Thickness = (borders.BottomSize?.Value ?? 4U) / 8D,
                Color = ParseNativeColor(NormalizeNativeBorderColor(borders.BottomColorHex)) ?? PdfCore.PdfColor.Black,
                SpacingBefore = borders.BottomSpace?.Value ?? 0D,
                SpacingAfter = paragraphStyle.SpacingAfter ?? 6D,
                KeepWithNext = paragraphStyle.KeepWithNext
            };
        }

        private static PdfCore.PdfHorizontalRuleStyle? CreateNativeTopBorderRuleStyle(WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle) {
            WordParagraphBorders borders = paragraph.Borders;
            if (!HasNativeOnlyTopParagraphBorder(borders)) {
                return null;
            }

            return new PdfCore.PdfHorizontalRuleStyle {
                Thickness = (borders.TopSize?.Value ?? 4U) / 8D,
                Color = ParseNativeColor(NormalizeNativeBorderColor(borders.TopColorHex)) ?? PdfCore.PdfColor.Black,
                SpacingBefore = paragraphStyle.SpacingBefore,
                SpacingAfter = borders.TopSpace?.Value ?? 0D,
                KeepWithNext = true
            };
        }

        private static (PdfCore.PdfColor? Color, double Width)? GetNativeUniformParagraphBorder(WordParagraphBorders borders) {
            if (!HasNativeBorder(borders.TopStyle) ||
                !HasNativeBorder(borders.BottomStyle) ||
                !HasNativeBorder(borders.LeftStyle) ||
                !HasNativeBorder(borders.RightStyle)) {
                return null;
            }

            if (borders.TopStyle != borders.BottomStyle ||
                borders.TopStyle != borders.LeftStyle ||
                borders.TopStyle != borders.RightStyle) {
                return null;
            }

            uint topSize = borders.TopSize?.Value ?? 4U;
            if (topSize != (borders.BottomSize?.Value ?? 4U) ||
                topSize != (borders.LeftSize?.Value ?? 4U) ||
                topSize != (borders.RightSize?.Value ?? 4U)) {
                return null;
            }

            string? topColor = NormalizeNativeBorderColor(borders.TopColorHex);
            if (!string.Equals(topColor, NormalizeNativeBorderColor(borders.BottomColorHex), StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(topColor, NormalizeNativeBorderColor(borders.LeftColorHex), StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(topColor, NormalizeNativeBorderColor(borders.RightColorHex), StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            PdfCore.PdfColor color = ParseNativeColor(topColor) ?? PdfCore.PdfColor.Black;
            return (color, topSize / 8D);
        }

        private static bool HasNativeBorder(W.BorderValues? style) =>
            style != null && style != W.BorderValues.Nil && style != W.BorderValues.None;

        private static bool HasNativeParagraphBorder(WordParagraphBorders borders) =>
            HasNativeBorder(borders.TopStyle) ||
            HasNativeBorder(borders.RightStyle) ||
            HasNativeBorder(borders.BottomStyle) ||
            HasNativeBorder(borders.LeftStyle);

        private static PdfCore.PdfPanelBorder? CreateNativePanelBorder(W.BorderValues? borderStyle, string? color, DocumentFormat.OpenXml.UInt32Value? size) {
            if (!HasNativeBorder(borderStyle)) {
                return null;
            }

            return new PdfCore.PdfPanelBorder {
                Color = ParseNativeColor(NormalizeNativeBorderColor(color)) ?? PdfCore.PdfColor.Black,
                Width = (size?.Value ?? 4U) / 8D
            };
        }

        private static bool HasNativeOnlyBottomParagraphBorder(WordParagraphBorders borders) =>
            HasNativeBorder(borders.BottomStyle) &&
            !HasNativeBorder(borders.TopStyle) &&
            !HasNativeBorder(borders.LeftStyle) &&
            !HasNativeBorder(borders.RightStyle);

        private static bool HasNativeOnlyTopParagraphBorder(WordParagraphBorders borders) =>
            HasNativeBorder(borders.TopStyle) &&
            !HasNativeBorder(borders.BottomStyle) &&
            !HasNativeBorder(borders.LeftStyle) &&
            !HasNativeBorder(borders.RightStyle);

        private static string? NormalizeNativeBorderColor(string? color) =>
            string.IsNullOrWhiteSpace(color) || string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase)
                ? null
                : color;

        private static PdfCore.PdfAlign MapNativeParagraphAlign(W.JustificationValues? alignment, bool allowJustify = true) {
            if (alignment == W.JustificationValues.Center) {
                return PdfCore.PdfAlign.Center;
            }

            if (alignment == W.JustificationValues.Right) {
                return PdfCore.PdfAlign.Right;
            }

            if (allowJustify &&
                (alignment == W.JustificationValues.Both ||
                 alignment == W.JustificationValues.Distribute ||
                 alignment == W.JustificationValues.HighKashida ||
                 alignment == W.JustificationValues.LowKashida ||
                 alignment == W.JustificationValues.MediumKashida ||
                 alignment == W.JustificationValues.ThaiDistribute)) {
                return PdfCore.PdfAlign.Justify;
            }

            return PdfCore.PdfAlign.Left;
        }

        private static PdfCore.PdfColumnAlign MapNativeColumnAlign(W.JustificationValues? alignment) {
            if (alignment == W.JustificationValues.Center) {
                return PdfCore.PdfColumnAlign.Center;
            }

            if (alignment == W.JustificationValues.Right) {
                return PdfCore.PdfColumnAlign.Right;
            }

            return PdfCore.PdfColumnAlign.Left;
        }

        private static PdfCore.PdfCellVerticalAlign MapNativeCellVerticalAlign(W.TableVerticalAlignmentValues? alignment) {
            if (alignment == W.TableVerticalAlignmentValues.Center) {
                return PdfCore.PdfCellVerticalAlign.Middle;
            }

            if (alignment == W.TableVerticalAlignmentValues.Bottom) {
                return PdfCore.PdfCellVerticalAlign.Bottom;
            }

            return PdfCore.PdfCellVerticalAlign.Top;
        }

        private static int GetHeadingLevel(WordParagraph paragraph) {
            if (!paragraph.Style.HasValue) {
                return 0;
            }

            return paragraph.Style.Value switch {
                WordParagraphStyles.Heading1 => 1,
                WordParagraphStyles.Heading2 => 2,
                WordParagraphStyles.Heading3 => 3,
                WordParagraphStyles.Heading4 => 3,
                WordParagraphStyles.Heading5 => 3,
                WordParagraphStyles.Heading6 => 3,
                _ => 0
            };
        }

        private static PdfCore.PdfColor? GetNativeHeadingColor(int headingLevel, PdfCore.PdfColor? explicitColor) {
            if (explicitColor.HasValue || headingLevel <= 0) {
                return explicitColor;
            }

            return PdfCore.PdfColor.FromRgb(47, 84, 150);
        }

        private static PdfCore.PageSize GetNativePageSize(WordSection section, PdfSaveOptions? options) {
            PdfCore.PageSize size;
            if (options?.PageSize != null) {
                size = options.PageSize.Value;
                if (options.Orientation == null) {
                    return size;
                }
            } else if (section.PageSettings.Width?.Value > 0 && section.PageSettings.Height?.Value > 0) {
                size = new PdfCore.PageSize(section.PageSettings.Width.Value / 20D, section.PageSettings.Height.Value / 20D);
            } else if (section.PageSettings.PageSize.HasValue) {
                size = MapNativePageSize(section.PageSettings.PageSize.Value);
            } else if (options?.DefaultPageSize.HasValue == true) {
                size = MapNativePageSize(options.DefaultPageSize.Value);
            } else {
                size = PdfCore.PageSizes.A4;
            }

            PdfPageOrientation orientation;
            if (options?.Orientation != null) {
                orientation = options.Orientation.Value;
            } else if (section.PageSettings.Orientation == W.PageOrientationValues.Landscape) {
                orientation = PdfPageOrientation.Landscape;
            } else if (options?.DefaultOrientation != null) {
                orientation = options.DefaultOrientation == W.PageOrientationValues.Landscape ? PdfPageOrientation.Landscape : PdfPageOrientation.Portrait;
            } else {
                orientation = PdfPageOrientation.Portrait;
            }

            return orientation == PdfPageOrientation.Landscape ? size.Landscape() : size.Portrait();
        }

        private static PdfCore.PageSize MapNativePageSize(WordPageSize pageSize) =>
            pageSize switch {
                WordPageSize.Letter => PdfCore.PageSizes.Letter,
                WordPageSize.Legal => PdfCore.PageSizes.Legal,
                WordPageSize.A3 => new PdfCore.PageSize(842, 1191),
                WordPageSize.A4 => PdfCore.PageSizes.A4,
                WordPageSize.A5 => PdfCore.PageSizes.A5,
                WordPageSize.A6 => new PdfCore.PageSize(298, 420),
                WordPageSize.B5 => new PdfCore.PageSize(499, 709),
                WordPageSize.Executive => new PdfCore.PageSize(522, 756),
                WordPageSize.Statement => new PdfCore.PageSize(396, 612),
                _ => PdfCore.PageSizes.A4
            };

        private static PdfCore.PageMargins GetNativeMargins(WordSection section, PdfSaveOptions? options) {
            if (options?.Margins != null) {
                return options.Margins.Value;
            }

            return new PdfCore.PageMargins(
                (section.Margins.Left?.Value ?? 0) / 20D,
                (section.Margins.Top ?? 0) / 20D,
                (section.Margins.Right?.Value ?? 0) / 20D,
                (section.Margins.Bottom ?? 0) / 20D);
        }

        private static string GetNativePageNumberFormat(PdfSaveOptions? options) {
            string? format = options?.PageNumberFormat;
            if (string.IsNullOrWhiteSpace(format)) {
                return "{page}/{pages}";
            }

            return format!.Replace("{current}", "{page}").Replace("{total}", "{pages}");
        }

        private static string? BuildNativeKeywords(PdfSaveOptions? options, BuiltinDocumentProperties properties) {
            return options?.Keywords ?? properties.Keywords;
        }

        private static PdfCore.PdfColor? ParseNativeColor(string? hex) {
            if (hex == null || string.IsNullOrWhiteSpace(hex) || hex.Equals("auto", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            string value = hex.Trim();
            if (value.StartsWith("#", StringComparison.Ordinal)) {
                value = value.Substring(1);
            }

            if (value.Length != 6 ||
                !byte.TryParse(value.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte r) ||
                !byte.TryParse(value.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte g) ||
                !byte.TryParse(value.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte b)) {
                return null;
            }

            return PdfCore.PdfColor.FromRgb(r, g, b);
        }

        private static PdfCore.PdfColor? MapNativeHighlight(W.HighlightColorValues? highlight) {
            if (!highlight.HasValue || highlight.Value == W.HighlightColorValues.None) {
                return null;
            }

            if (highlight.Value == W.HighlightColorValues.Black) return PdfCore.PdfColor.Black;
            if (highlight.Value == W.HighlightColorValues.Blue) return PdfCore.PdfColor.FromRgb(0, 0, 255);
            if (highlight.Value == W.HighlightColorValues.Cyan) return PdfCore.PdfColor.FromRgb(0, 255, 255);
            if (highlight.Value == W.HighlightColorValues.Green) return PdfCore.PdfColor.FromRgb(0, 255, 0);
            if (highlight.Value == W.HighlightColorValues.Magenta) return PdfCore.PdfColor.FromRgb(255, 0, 255);
            if (highlight.Value == W.HighlightColorValues.Red) return PdfCore.PdfColor.FromRgb(255, 0, 0);
            if (highlight.Value == W.HighlightColorValues.Yellow) return PdfCore.PdfColor.FromRgb(255, 255, 0);
            if (highlight.Value == W.HighlightColorValues.White) return PdfCore.PdfColor.White;
            if (highlight.Value == W.HighlightColorValues.DarkBlue) return PdfCore.PdfColor.FromRgb(0, 0, 139);
            if (highlight.Value == W.HighlightColorValues.DarkCyan) return PdfCore.PdfColor.FromRgb(0, 139, 139);
            if (highlight.Value == W.HighlightColorValues.DarkGreen) return PdfCore.PdfColor.FromRgb(0, 100, 0);
            if (highlight.Value == W.HighlightColorValues.DarkMagenta) return PdfCore.PdfColor.FromRgb(139, 0, 139);
            if (highlight.Value == W.HighlightColorValues.DarkRed) return PdfCore.PdfColor.FromRgb(139, 0, 0);
            if (highlight.Value == W.HighlightColorValues.DarkYellow) return PdfCore.PdfColor.FromRgb(184, 134, 11);
            if (highlight.Value == W.HighlightColorValues.LightGray) return PdfCore.PdfColor.FromRgb(211, 211, 211);
            if (highlight.Value == W.HighlightColorValues.DarkGray) return PdfCore.PdfColor.FromRgb(169, 169, 169);

            return null;
        }
    }
}
