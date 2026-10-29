using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const double MaxNativeParagraphBorderSpacingPoints = 31D;
        // Word places side borders outside the text frame, including a 1.5-point gap.
        private const double NativeParagraphBorderHorizontalGapPoints = 1.5D;

        private static PdfCore.PdfPanelStyle? CreateNativeParagraphPanelStyle(WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle) {
            NativeParagraphBorders borders = GetNativeEffectiveParagraphBorders(paragraph);
            global::OfficeIMO.OfficeColor? background = ParseNativeColor(GetNativeEffectiveParagraphShadingFill(paragraph));
            (global::OfficeIMO.OfficeColor? Color, double Width)? border = GetNativeUniformParagraphBorder(borders);
            bool renderAsRule = !background.HasValue && !HasNativeBorder(borders.Between.Style) &&
                (HasNativeOnlyTopParagraphBorder(borders) || HasNativeOnlyBottomParagraphBorder(borders));
            if (renderAsRule) {
                return null;
            }

            bool hasParagraphBorder = HasNativeParagraphBorder(borders);
            if (!background.HasValue && border == null && !hasParagraphBorder) {
                return null;
            }

            bool backgroundOnly = background.HasValue && border == null && !hasParagraphBorder;
            var style = new PdfCore.PdfPanelStyle {
                Background = background,
                BorderColor = border?.Color,
                BorderWidth = border?.Width ?? 0D,
                PaddingX = 0D,
                PaddingY = 0D,
                PaddingTopOverride = backgroundOnly ? 0D : ResolveNativeParagraphPanelPadding(borders.Top),
                PaddingBottomOverride = backgroundOnly ? 0D : ResolveNativeParagraphPanelPadding(borders.Bottom),
                RepeatFragmentDecoration = !hasParagraphBorder,
                AllowTrailingPaddingOverflow = hasParagraphBorder,
                FragmentBottomInset = hasParagraphBorder && !UsesModernNativeWordLayout(paragraph._document)
                    ? ResolveNativeParagraphPanelPadding(borders.Bottom) : 0D,
                SpacingBefore = paragraphStyle.SpacingBefore,
                SpacingAfter = paragraphStyle.SpacingAfter ?? 6D,
                ClipSpacingAfterAtFrameEnd = true,
                KeepWithNext = paragraphStyle.KeepWithNext,
                Align = ResolveNativeParagraphAlign(paragraph, allowJustify: false)
            };

            if (hasParagraphBorder) {
                style.TopBorder = CreateNativePanelBorder(borders.Top);
                style.RightBorder = CreateNativePanelBorder(borders.Right);
                style.BottomBorder = CreateNativePanelBorder(borders.Bottom);
                style.LeftBorder = CreateNativePanelBorder(borders.Left);
                if (style.TopBorder is { } top) {
                    top.Offset = -top.Width / 2D;
                    style.TopBorder = top;
                }
                if (style.BottomBorder is { } bottom) {
                    bottom.Offset = -bottom.Width / 2D;
                    style.BottomBorder = bottom;
                }
                if (style.LeftBorder is { } left) {
                    left.Offset = NativeParagraphBorderHorizontalGapPoints + Math.Min(borders.Left.Space ?? 0D, MaxNativeParagraphBorderSpacingPoints) + left.Width / 2D;
                    style.LeftBorder = left;
                }
                if (style.RightBorder is { } right) {
                    right.Offset = NativeParagraphBorderHorizontalGapPoints + Math.Min(borders.Right.Space ?? 0D, MaxNativeParagraphBorderSpacingPoints) + right.Width / 2D;
                    style.RightBorder = right;
                }
            }

            JoinNativeParagraphBorderGroup(paragraph, borders, style);
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

            return HasNativeOnlyBottomParagraphBorder(GetNativeEffectiveParagraphBorders(paragraph));
        }

        private static PdfCore.PdfHorizontalRuleStyle? CreateNativeHorizontalRuleStyle(WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle) {
            NativeParagraphBorders borders = GetNativeEffectiveParagraphBorders(paragraph);
            if (!HasNativeBorder(borders.Bottom.Style)) {
                return null;
            }

            return new PdfCore.PdfHorizontalRuleStyle {
                Thickness = (borders.Bottom.Size ?? 4U) / 8D,
                Color = ParseNativeColor(NormalizeNativeBorderColor(borders.Bottom.ColorHex)) ?? global::OfficeIMO.OfficeColor.Black,
                SpacingBefore = paragraphStyle.SpacingBefore,
                SpacingAfter = paragraphStyle.SpacingAfter ?? (borders.Bottom.Space ?? 6U),
                KeepWithNext = paragraphStyle.KeepWithNext
            };
        }

        private static PdfCore.PdfHorizontalRuleStyle? CreateNativeBottomBorderRuleStyle(WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle) {
            NativeParagraphBorders borders = GetNativeEffectiveParagraphBorders(paragraph);
            if (!HasNativeOnlyBottomParagraphBorder(borders) || HasNativeMatchingBorderNeighbor(paragraph, borders, next: true)) {
                return null;
            }

            return new PdfCore.PdfHorizontalRuleStyle {
                Thickness = (borders.Bottom.Size ?? 4U) / 8D,
                Color = ParseNativeColor(NormalizeNativeBorderColor(borders.Bottom.ColorHex)) ?? global::OfficeIMO.OfficeColor.Black,
                SpacingBefore = borders.Bottom.Space ?? 0D,
                SpacingAfter = paragraphStyle.SpacingAfter ?? 6D,
                KeepWithNext = paragraphStyle.KeepWithNext
            };
        }

        private static PdfCore.PdfHorizontalRuleStyle? CreateNativeTopBorderRuleStyle(WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle) {
            NativeParagraphBorders borders = GetNativeEffectiveParagraphBorders(paragraph);
            if (!HasNativeOnlyTopParagraphBorder(borders) || HasNativeMatchingBorderNeighbor(paragraph, borders, next: false)) {
                return null;
            }

            return new PdfCore.PdfHorizontalRuleStyle {
                Thickness = (borders.Top.Size ?? 4U) / 8D,
                Color = ParseNativeColor(NormalizeNativeBorderColor(borders.Top.ColorHex)) ?? global::OfficeIMO.OfficeColor.Black,
                SpacingBefore = paragraphStyle.SpacingBefore,
                SpacingAfter = borders.Top.Space ?? 0D,
                KeepWithNext = true
            };
        }

        private static (global::OfficeIMO.OfficeColor? Color, double Width)? GetNativeUniformParagraphBorder(NativeParagraphBorders borders) {
            if (!HasNativeBorder(borders.Top.Style) ||
                !HasNativeBorder(borders.Bottom.Style) ||
                !HasNativeBorder(borders.Left.Style) ||
                !HasNativeBorder(borders.Right.Style)) {
                return null;
            }

            if (borders.Top.Style != borders.Bottom.Style ||
                borders.Top.Style != borders.Left.Style ||
                borders.Top.Style != borders.Right.Style) {
                return null;
            }

            uint topSize = borders.Top.Size ?? 4U;
            if (topSize != (borders.Bottom.Size ?? 4U) ||
                topSize != (borders.Left.Size ?? 4U) ||
                topSize != (borders.Right.Size ?? 4U)) {
                return null;
            }

            string? topColor = NormalizeNativeBorderColor(borders.Top.ColorHex);
            if (!string.Equals(topColor, NormalizeNativeBorderColor(borders.Bottom.ColorHex), StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(topColor, NormalizeNativeBorderColor(borders.Left.ColorHex), StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(topColor, NormalizeNativeBorderColor(borders.Right.ColorHex), StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            global::OfficeIMO.OfficeColor color = ParseNativeColor(topColor) ?? global::OfficeIMO.OfficeColor.Black;
            return (color, topSize / 8D);
        }

        private static bool HasNativeBorder(W.BorderValues? style) =>
            style != null && style != W.BorderValues.Nil && style != W.BorderValues.None;

        private static bool HasNativeBorder(WordBorderStyle? style) =>
            HasNativeBorder(style.ToOpenXml());

        private static bool HasNativeParagraphBorder(NativeParagraphBorders borders) =>
            HasNativeBorder(borders.Top.Style) ||
            HasNativeBorder(borders.Right.Style) ||
            HasNativeBorder(borders.Bottom.Style) ||
            HasNativeBorder(borders.Left.Style) ||
            HasNativeBorder(borders.Between.Style);

        private static PdfCore.PdfPanelBorder? CreateNativePanelBorder(NativeParagraphBorderSide border) {
            if (!HasNativeBorder(border.Style)) {
                return null;
            }

            return new PdfCore.PdfPanelBorder {
                Color = ParseNativeColor(NormalizeNativeBorderColor(border.ColorHex)) ?? global::OfficeIMO.OfficeColor.Black,
                Width = (border.Size ?? 4U) / 8D
            };
        }

        private static double ResolveNativeParagraphPanelPadding(NativeParagraphBorderSide border) =>
            HasNativeBorder(border.Style)
                ? Math.Min(border.Space ?? 0D, MaxNativeParagraphBorderSpacingPoints) + (border.Size ?? 4U) / 8D
                : 0D;

        private static bool HasNativeOnlyBottomParagraphBorder(NativeParagraphBorders borders) =>
            HasNativeBorder(borders.Bottom.Style) &&
            !HasNativeBorder(borders.Between.Style) &&
            !HasNativeBorder(borders.Top.Style) &&
            !HasNativeBorder(borders.Left.Style) &&
            !HasNativeBorder(borders.Right.Style);

        private static bool HasNativeOnlyTopParagraphBorder(NativeParagraphBorders borders) =>
            HasNativeBorder(borders.Top.Style) &&
            !HasNativeBorder(borders.Between.Style) &&
            !HasNativeBorder(borders.Bottom.Style) &&
            !HasNativeBorder(borders.Left.Style) &&
            !HasNativeBorder(borders.Right.Style);

        private static string? NormalizeNativeBorderColor(string? color) =>
            string.IsNullOrWhiteSpace(color) || string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase)
                ? null
                : color;

        private static string? NormalizeNativeShadingFill(string? color) =>
            string.IsNullOrWhiteSpace(color) || string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase)
                ? null
                : color;

        private static string? GetNativeEffectiveParagraphShadingFill(WordParagraph paragraph) =>
            NormalizeNativeShadingFill(paragraph.ShadingFillColorHex) ?? GetNativeParagraphStyleDefaults(paragraph).ShadingFillColorHex;

        private static NativeParagraphBorders GetNativeEffectiveParagraphBorders(WordParagraph paragraph) =>
            MergeNativeParagraphBorders(GetNativeParagraphStyleDefaults(paragraph).Borders, GetNativeDirectParagraphBorders(paragraph));

        private static NativeParagraphBorders GetNativeDirectParagraphBorders(WordParagraph paragraph) {
            W.ParagraphBorders? borders = paragraph._paragraph?.ParagraphProperties?.ParagraphBorders;
            return new NativeParagraphBorders(
                ReadNativeParagraphBorderSide(borders?.TopBorder) ?? default,
                ReadNativeParagraphBorderSide(borders?.RightBorder) ?? default,
                ReadNativeParagraphBorderSide(borders?.BottomBorder) ?? default,
                ReadNativeParagraphBorderSide(borders?.LeftBorder) ?? default) {
                Between = ReadNativeParagraphBorderSide(borders?.BetweenBorder) ?? default
            };
        }

        private static NativeParagraphBorders MergeNativeParagraphBorders(NativeParagraphBorders styleBorders, NativeParagraphBorders directBorders) =>
            styleBorders with {
                Top = directBorders.Top.IsEmpty ? styleBorders.Top : directBorders.Top,
                Right = directBorders.Right.IsEmpty ? styleBorders.Right : directBorders.Right,
                Bottom = directBorders.Bottom.IsEmpty ? styleBorders.Bottom : directBorders.Bottom,
                Left = directBorders.Left.IsEmpty ? styleBorders.Left : directBorders.Left,
                Between = directBorders.Between.IsEmpty ? styleBorders.Between : directBorders.Between
            };

    }
}
