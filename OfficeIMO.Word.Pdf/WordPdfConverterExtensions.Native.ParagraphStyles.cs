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
        private const double MaxNativeParagraphBorderSpacingPoints = 31D;

        private static PdfCore.PdfParagraphStyle CreateNativeParagraphStyle(WordParagraph paragraph) =>
            CreateNativeParagraphStyle(paragraph, GetNativeDocumentDefaults(paragraph._document));

        private static PdfCore.PdfParagraphStyle CreateNativeParagraphStyle(
            WordParagraph paragraph,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap? nativeFontMap = null) {
            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(paragraph);
            var style = new PdfCore.PdfParagraphStyle();
            double fontSize = ResolveNativeParagraphEffectiveFontSize(paragraph, nativeDefaults, styleDefaults);
            double lineHeight = ResolveNativeParagraphLineHeight(
                paragraph,
                fontSize,
                nativeDefaults,
                styleDefaults,
                nativeFontMap);
            W.SpacingBetweenLines? directSpacing = paragraph._paragraph?.ParagraphProperties?.GetFirstChild<W.SpacingBetweenLines>();
            if (paragraph.LineSpacingBeforePoints.HasValue) {
                style.SpacingBefore = paragraph.LineSpacingBeforePoints.Value;
            } else if (GetNativeSpacingBeforePoints(directSpacing, fontSize, lineHeight) is { } directSpacingBefore) {
                style.SpacingBefore = directSpacingBefore;
            } else if (styleDefaults.SpacingBefore.HasValue) {
                style.SpacingBefore = styleDefaults.SpacingBefore.Value;
            } else if (nativeDefaults.ParagraphSpacingBeforeDeclared) {
                style.SpacingBefore = nativeDefaults.ParagraphSpacingBefore;
            }

            if (paragraph.LineSpacingAfterPoints.HasValue) {
                style.SpacingAfter = paragraph.LineSpacingAfterPoints.Value;
            } else if (GetNativeSpacingAfterPoints(directSpacing, fontSize, lineHeight) is { } directSpacingAfter) {
                style.SpacingAfter = directSpacingAfter;
            } else if (styleDefaults.SpacingAfter.HasValue) {
                style.SpacingAfter = styleDefaults.SpacingAfter.Value;
            }

            if (paragraph.IndentationBeforePoints.HasValue) {
                style.LeftIndent = paragraph.IndentationBeforePoints.Value;
            } else if (styleDefaults.LeftIndent.HasValue) {
                style.LeftIndent = styleDefaults.LeftIndent.Value;
            }

            if (paragraph.IndentationAfterPoints.HasValue) {
                style.RightIndent = paragraph.IndentationAfterPoints.Value;
            } else if (styleDefaults.RightIndent.HasValue) {
                style.RightIndent = styleDefaults.RightIndent.Value;
            }

            if (paragraph.IndentationFirstLinePoints.HasValue) {
                style.FirstLineIndent = paragraph.IndentationFirstLinePoints.Value;
            } else if (styleDefaults.FirstLineIndent.HasValue) {
                style.FirstLineIndent = styleDefaults.FirstLineIndent.Value;
            }

            if (paragraph.IndentationHangingPoints.HasValue) {
                double hangingIndent = paragraph.IndentationHangingPoints.Value;
                if (style.LeftIndent < hangingIndent) {
                    style.LeftIndent = hangingIndent;
                }

                style.FirstLineIndent = -hangingIndent;
            } else if (style.FirstLineIndent < 0D && style.LeftIndent < -style.FirstLineIndent) {
                style.LeftIndent = -style.FirstLineIndent;
            }

            style.LineHeight = lineHeight;
            if (nativeDefaults.DefaultTabStopWidth.HasValue) {
                style.DefaultTabStopWidth = nativeDefaults.DefaultTabStopWidth.Value;
            }

            if (!paragraph.LineSpacingAfterPoints.HasValue &&
                GetNativeSpacingAfterPoints(directSpacing, fontSize, lineHeight) == null &&
                !styleDefaults.SpacingAfter.HasValue) {
                style.SpacingAfter = nativeDefaults.ParagraphSpacingAfter;
            }

            foreach (WordTabStop tabStop in GetNativeParagraphEffectiveTabStops(paragraph)
                .Where(tabStop => tabStop.Position > 0 && IsNativeRenderableTextTabStop(tabStop.Alignment))
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

            style.KeepTogether = ReadNativeDirectParagraphOnOff<W.KeepLines>(paragraph) ?? styleDefaults.KeepTogether ?? false;
            style.KeepWithNext = ReadNativeDirectParagraphOnOff<W.KeepNext>(paragraph) ?? styleDefaults.KeepWithNext ?? false;
            style.WidowControl = ReadNativeDirectParagraphOnOff<W.WidowControl>(paragraph) ?? styleDefaults.WidowControl ?? nativeDefaults.ParagraphWidowControl;
            return style;
        }

        private static bool ShouldSuppressNativeContextualSpacingAfter(WordParagraph paragraph, WordParagraph? nextParagraph) {
            if (nextParagraph == null ||
                paragraph.IsPageBreak ||
                nextParagraph.IsPageBreak ||
                HasNativePageBreakBefore(nextParagraph)) {
                return false;
            }

            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(paragraph);
            bool contextualSpacing = ReadNativeDirectParagraphOnOff<W.ContextualSpacing>(paragraph) ?? styleDefaults.ContextualSpacing ?? false;
            return contextualSpacing &&
                   string.Equals(
                       GetNativeParagraphStyleIdentity(paragraph),
                       GetNativeParagraphStyleIdentity(nextParagraph),
                       StringComparison.Ordinal);
        }

        private static string GetNativeParagraphStyleIdentity(WordParagraph paragraph) {
            IReadOnlyList<W.Style> styleChain = GetNativeParagraphStyleChain(paragraph._document, paragraph.StyleId);
            if (styleChain.Count > 0) {
                return styleChain[styleChain.Count - 1].StyleId?.Value ?? string.Empty;
            }

            return paragraph.StyleId ?? string.Empty;
        }

        private static bool IsNativeRenderableTextTabStop(W.TabStopValues alignment) =>
            alignment != W.TabStopValues.Bar &&
            alignment != W.TabStopValues.Clear;

        private static double ResolveNativeParagraphLineHeight(WordParagraph paragraph, double fontSize) =>
            ResolveNativeParagraphLineHeight(paragraph, fontSize, GetNativeDocumentDefaults(paragraph._document));

        private static double ResolveNativeParagraphLineHeight(WordParagraph paragraph, double fontSize, NativeDocumentDefaults nativeDefaults) {
            return ResolveNativeParagraphLineHeight(paragraph, fontSize, nativeDefaults, GetNativeParagraphStyleDefaults(paragraph));
        }

        private static double ResolveNativeParagraphFontSize(WordParagraph paragraph, NativeDocumentDefaults nativeDefaults, NativeParagraphStyleDefaults styleDefaults) =>
            paragraph.FontSize.HasValue && paragraph.FontSize.Value > 0
                ? paragraph.FontSize.Value
                : styleDefaults.FontSize ?? nativeDefaults.FontSize;

        private static double ResolveNativeParagraphEffectiveFontSize(WordParagraph paragraph, NativeDocumentDefaults nativeDefaults, NativeParagraphStyleDefaults styleDefaults) =>
            ResolveNativeParagraphEffectiveFontSize(paragraph, nativeDefaults, styleDefaults, NativeTableRunStyleDefaults.Empty);

        private static double ResolveNativeParagraphEffectiveFontSize(WordParagraph paragraph, NativeDocumentDefaults nativeDefaults, NativeParagraphStyleDefaults styleDefaults, NativeTableRunStyleDefaults tableRunStyleDefaults) {
            double fontSize = ResolveNativeParagraphFontSize(paragraph, nativeDefaults, styleDefaults);
            List<WordParagraph> runs = GetNativeRuns(paragraph);
            if (runs.Count == 0 && !string.IsNullOrWhiteSpace(paragraph.Text)) {
                NativeResolvedTextStyle paragraphTextStyle = ResolveNativeTextRunStyle(paragraph, tableRunStyleDefaults: tableRunStyleDefaults, nativeDefaults: nativeDefaults);
                if (paragraphTextStyle.FontSize.HasValue && paragraphTextStyle.FontSize.Value > fontSize) {
                    fontSize = paragraphTextStyle.FontSize.Value;
                }
            }

            foreach (WordParagraph run in runs) {
                if (run.IsImage || string.IsNullOrWhiteSpace(run.Text)) {
                    continue;
                }

                NativeResolvedTextStyle runTextStyle = ResolveNativeTextRunStyle(run, paragraph, tableRunStyleDefaults, nativeDefaults);
                if (runTextStyle.FontSize.HasValue && runTextStyle.FontSize.Value > fontSize) {
                    fontSize = runTextStyle.FontSize.Value;
                }
            }

            return fontSize;
        }

        private static double ResolveNativeParagraphLineHeight(
            WordParagraph paragraph,
            double fontSize,
            NativeDocumentDefaults nativeDefaults,
            NativeParagraphStyleDefaults styleDefaults,
            NativeFontMap? nativeFontMap = null) {
            double naturalLineHeight = ResolveNativeParagraphSingleLineHeight(
                paragraph,
                nativeDefaults,
                styleDefaults,
                nativeFontMap: nativeFontMap);
            if (paragraph.LineSpacing.HasValue && paragraph.LineSpacingRule == W.LineSpacingRuleValues.Auto) {
                return Math.Max(0.01D, naturalLineHeight * (paragraph.LineSpacing.Value / 240D));
            }

            if (paragraph.LineSpacingPoints.HasValue && fontSize > 0D) {
                return ResolveNativeLineSpacingHeight(paragraph.LineSpacingPoints.Value, paragraph.LineSpacingRule, fontSize, naturalLineHeight);
            }

            if (styleDefaults.LineSpacingPoints.HasValue && fontSize > 0D) {
                return ResolveNativeLineSpacingHeight(styleDefaults.LineSpacingPoints.Value, styleDefaults.LineSpacingRule, fontSize, naturalLineHeight);
            }

            if (styleDefaults.LineHeight.HasValue) {
                return styleDefaults.LineHeight.Value;
            }

            return nativeDefaults.ParagraphLineHeight;
        }

        private static double ResolveNativeParagraphSingleLineHeight(
            WordParagraph paragraph,
            NativeDocumentDefaults nativeDefaults,
            NativeParagraphStyleDefaults styleDefaults,
            NativeTableRunStyleDefaults tableRunStyleDefaults = default,
            NativeFontMap? nativeFontMap = null) {
            double lineHeight = ResolveNativeWordSingleLineHeight(
                nativeFontMap,
                paragraph.FontFamily,
                paragraph.FontFamilyHighAnsi,
                paragraph.FontFamilyEastAsia,
                paragraph.FontFamilyComplexScript,
                styleDefaults.FontFamily,
                tableRunStyleDefaults.FontFamily,
                nativeDefaults.FontFamily);
            foreach (WordParagraph run in GetNativeRuns(paragraph)) {
                if (run.IsImage || string.IsNullOrWhiteSpace(run.Text)) {
                    continue;
                }

                NativeCharacterStyleDefaults characterStyle =
                    GetNativeCharacterStyleDefaults(run._document, GetNativeRunProperties(run));
                lineHeight = Math.Max(
                    lineHeight,
                    ResolveNativeWordSingleLineHeight(
                        nativeFontMap,
                        run.FontFamily,
                        run.FontFamilyHighAnsi,
                        run.FontFamilyEastAsia,
                        run.FontFamilyComplexScript,
                        characterStyle.FontFamily,
                        styleDefaults.FontFamily,
                        tableRunStyleDefaults.FontFamily,
                        nativeDefaults.FontFamily));
            }

            return lineHeight;
        }

        private static double ResolveNativeLineSpacingHeight(double lineSpacingPoints, W.LineSpacingRuleValues? lineSpacingRule, double fontSize, double naturalLineHeight) {
            double requestedLineHeight = lineSpacingPoints / fontSize;
            if (lineSpacingRule == W.LineSpacingRuleValues.AtLeast) {
                return Math.Max(naturalLineHeight, requestedLineHeight);
            }

            return requestedLineHeight;
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
            NativeParagraphBorders borders = GetNativeEffectiveParagraphBorders(paragraph);
            PdfCore.PdfColor? background = ParseNativeColor(GetNativeEffectiveParagraphShadingFill(paragraph));
            (PdfCore.PdfColor? Color, double Width)? border = GetNativeUniformParagraphBorder(borders);
            bool renderAsRule = !background.HasValue &&
                (HasNativeOnlyTopParagraphBorder(borders) || HasNativeOnlyBottomParagraphBorder(borders));
            if (renderAsRule) {
                return null;
            }

            bool hasParagraphBorder = HasNativeParagraphBorder(borders);
            if (!background.HasValue && border == null && !hasParagraphBorder) {
                return null;
            }

            bool backgroundOnly = background.HasValue && border == null && !hasParagraphBorder;
            var style = new PdfCore.PanelStyle {
                Background = background,
                BorderColor = border?.Color,
                BorderWidth = border?.Width ?? 0D,
                PaddingX = ResolveNativeParagraphPanelPaddingX(borders, 6D),
                PaddingY = backgroundOnly ? 0D : ResolveNativeParagraphPanelPaddingY(borders, 4D),
                SpacingBefore = paragraphStyle.SpacingBefore,
                SpacingAfter = backgroundOnly ? 0D : paragraphStyle.SpacingAfter ?? 6D,
                Align = ResolveNativeParagraphAlign(paragraph, allowJustify: false)
            };

            if (border == null && hasParagraphBorder) {
                style.TopBorder = CreateNativePanelBorder(borders.Top);
                style.RightBorder = CreateNativePanelBorder(borders.Right);
                style.BottomBorder = CreateNativePanelBorder(borders.Bottom);
                style.LeftBorder = CreateNativePanelBorder(borders.Left);
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

            return HasNativeOnlyBottomParagraphBorder(GetNativeEffectiveParagraphBorders(paragraph));
        }

        private static PdfCore.PdfHorizontalRuleStyle? CreateNativeHorizontalRuleStyle(WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle) {
            NativeParagraphBorders borders = GetNativeEffectiveParagraphBorders(paragraph);
            if (!HasNativeBorder(borders.Bottom.Style)) {
                return null;
            }

            return new PdfCore.PdfHorizontalRuleStyle {
                Thickness = (borders.Bottom.Size ?? 4U) / 8D,
                Color = ParseNativeColor(NormalizeNativeBorderColor(borders.Bottom.ColorHex)) ?? PdfCore.PdfColor.Black,
                SpacingBefore = paragraphStyle.SpacingBefore,
                SpacingAfter = paragraphStyle.SpacingAfter ?? (borders.Bottom.Space ?? 6U),
                KeepWithNext = paragraphStyle.KeepWithNext
            };
        }

        private static PdfCore.PdfHorizontalRuleStyle? CreateNativeBottomBorderRuleStyle(WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle) {
            NativeParagraphBorders borders = GetNativeEffectiveParagraphBorders(paragraph);
            if (!HasNativeOnlyBottomParagraphBorder(borders)) {
                return null;
            }

            return new PdfCore.PdfHorizontalRuleStyle {
                Thickness = (borders.Bottom.Size ?? 4U) / 8D,
                Color = ParseNativeColor(NormalizeNativeBorderColor(borders.Bottom.ColorHex)) ?? PdfCore.PdfColor.Black,
                SpacingBefore = borders.Bottom.Space ?? 0D,
                SpacingAfter = paragraphStyle.SpacingAfter ?? 6D,
                KeepWithNext = paragraphStyle.KeepWithNext
            };
        }

        private static PdfCore.PdfHorizontalRuleStyle? CreateNativeTopBorderRuleStyle(WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle) {
            NativeParagraphBorders borders = GetNativeEffectiveParagraphBorders(paragraph);
            if (!HasNativeOnlyTopParagraphBorder(borders)) {
                return null;
            }

            return new PdfCore.PdfHorizontalRuleStyle {
                Thickness = (borders.Top.Size ?? 4U) / 8D,
                Color = ParseNativeColor(NormalizeNativeBorderColor(borders.Top.ColorHex)) ?? PdfCore.PdfColor.Black,
                SpacingBefore = paragraphStyle.SpacingBefore,
                SpacingAfter = borders.Top.Space ?? 0D,
                KeepWithNext = true
            };
        }

        private static (PdfCore.PdfColor? Color, double Width)? GetNativeUniformParagraphBorder(NativeParagraphBorders borders) {
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

            PdfCore.PdfColor color = ParseNativeColor(topColor) ?? PdfCore.PdfColor.Black;
            return (color, topSize / 8D);
        }

        private static bool HasNativeBorder(W.BorderValues? style) =>
            style != null && style != W.BorderValues.Nil && style != W.BorderValues.None;

        private static bool HasNativeParagraphBorder(NativeParagraphBorders borders) =>
            HasNativeBorder(borders.Top.Style) ||
            HasNativeBorder(borders.Right.Style) ||
            HasNativeBorder(borders.Bottom.Style) ||
            HasNativeBorder(borders.Left.Style);

        private static PdfCore.PdfPanelBorder? CreateNativePanelBorder(NativeParagraphBorderSide border) {
            if (!HasNativeBorder(border.Style)) {
                return null;
            }

            return new PdfCore.PdfPanelBorder {
                Color = ParseNativeColor(NormalizeNativeBorderColor(border.ColorHex)) ?? PdfCore.PdfColor.Black,
                Width = (border.Size ?? 4U) / 8D
            };
        }

        private static double ResolveNativeParagraphPanelPaddingX(NativeParagraphBorders borders, double defaultPadding) {
            uint? left = HasNativeBorder(borders.Left.Style) ? borders.Left.Space : null;
            uint? right = HasNativeBorder(borders.Right.Style) ? borders.Right.Space : null;
            if (!left.HasValue && !right.HasValue) {
                return defaultPadding;
            }

            return Math.Min(
                Math.Max(left.GetValueOrDefault(), right.GetValueOrDefault()),
                MaxNativeParagraphBorderSpacingPoints);
        }

        private static double ResolveNativeParagraphPanelPaddingY(NativeParagraphBorders borders, double defaultPadding) {
            uint? top = HasNativeBorder(borders.Top.Style) ? borders.Top.Space : null;
            uint? bottom = HasNativeBorder(borders.Bottom.Style) ? borders.Bottom.Space : null;
            if (!top.HasValue && !bottom.HasValue) {
                return defaultPadding;
            }

            return Math.Min(
                Math.Max(top.GetValueOrDefault(), bottom.GetValueOrDefault()),
                MaxNativeParagraphBorderSpacingPoints);
        }

        private static bool HasNativeOnlyBottomParagraphBorder(NativeParagraphBorders borders) =>
            HasNativeBorder(borders.Bottom.Style) &&
            !HasNativeBorder(borders.Top.Style) &&
            !HasNativeBorder(borders.Left.Style) &&
            !HasNativeBorder(borders.Right.Style);

        private static bool HasNativeOnlyTopParagraphBorder(NativeParagraphBorders borders) =>
            HasNativeBorder(borders.Top.Style) &&
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
            WordParagraphBorders borders = paragraph.Borders;
            return new NativeParagraphBorders(
                new NativeParagraphBorderSide(borders.TopStyle, NormalizeNativeBorderColor(borders.TopColorHex), borders.TopSize?.Value, borders.TopSpace?.Value),
                new NativeParagraphBorderSide(borders.RightStyle, NormalizeNativeBorderColor(borders.RightColorHex), borders.RightSize?.Value, borders.RightSpace?.Value),
                new NativeParagraphBorderSide(borders.BottomStyle, NormalizeNativeBorderColor(borders.BottomColorHex), borders.BottomSize?.Value, borders.BottomSpace?.Value),
                new NativeParagraphBorderSide(borders.LeftStyle, NormalizeNativeBorderColor(borders.LeftColorHex), borders.LeftSize?.Value, borders.LeftSpace?.Value));
        }

        private static NativeParagraphBorders MergeNativeParagraphBorders(NativeParagraphBorders styleBorders, NativeParagraphBorders directBorders) =>
            styleBorders with {
                Top = directBorders.Top.IsEmpty ? styleBorders.Top : directBorders.Top,
                Right = directBorders.Right.IsEmpty ? styleBorders.Right : directBorders.Right,
                Bottom = directBorders.Bottom.IsEmpty ? styleBorders.Bottom : directBorders.Bottom,
                Left = directBorders.Left.IsEmpty ? styleBorders.Left : directBorders.Left
            };

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

        private static W.JustificationValues? ResolveNativeParagraphJustification(WordParagraph paragraph) =>
            paragraph.ParagraphAlignment ?? GetNativeParagraphStyleDefaults(paragraph).Alignment;

        private static PdfCore.PdfAlign ResolveNativeParagraphAlign(WordParagraph paragraph, bool allowJustify = true) {
            W.JustificationValues? alignment = ResolveNativeParagraphJustification(paragraph);
            if (alignment == null && IsNativeBiDiParagraph(paragraph)) {
                return PdfCore.PdfAlign.Right;
            }

            return MapNativeParagraphAlign(alignment, allowJustify);
        }

        private static bool IsNativeBiDiParagraph(WordParagraph paragraph) =>
            paragraph.BiDi ||
            paragraph._paragraph?.ParagraphProperties?.GetFirstChild<W.BiDi>() != null;

        private static PdfCore.PdfColumnAlign ResolveNativeColumnAlign(WordParagraph paragraph) =>
            MapNativeColumnAlign(ResolveNativeParagraphJustification(paragraph));

        private static PdfCore.PdfColumnAlign MapNativeColumnAlign(W.JustificationValues? alignment) {
            if (alignment == W.JustificationValues.Center) {
                return PdfCore.PdfColumnAlign.Center;
            }

            if (alignment == W.JustificationValues.Right) {
                return PdfCore.PdfColumnAlign.Right;
            }

            return PdfCore.PdfColumnAlign.Left;
        }

        private static PdfCore.PdfCellVerticalAlign MapNativeCellVerticalAlign(W.TableVerticalAlignmentValues? alignment) =>
            MapNativeNullableCellVerticalAlign(alignment) ?? PdfCore.PdfCellVerticalAlign.Top;

        private static PdfCore.PdfCellVerticalAlign? MapNativeNullableCellVerticalAlign(W.TableVerticalAlignmentValues? alignment) {
            if (!alignment.HasValue) {
                return null;
            }

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

            PdfCore.PdfPageOrientation orientation;
            if (options?.Orientation != null) {
                orientation = options.Orientation.Value;
            } else if (section.PageSettings.Orientation == W.PageOrientationValues.Landscape) {
                orientation = PdfCore.PdfPageOrientation.Landscape;
            } else if (options?.DefaultOrientation != null) {
                orientation = options.DefaultOrientation == W.PageOrientationValues.Landscape ? PdfCore.PdfPageOrientation.Landscape : PdfCore.PdfPageOrientation.Portrait;
            } else {
                orientation = PdfCore.PdfPageOrientation.Portrait;
            }

            return orientation == PdfCore.PdfPageOrientation.Landscape ? size.Landscape() : size.Portrait();
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
            return GetNativeMargins(section, options, GetNativeHeaderFooterMarginExpansion(section, options));
        }

        private static PdfCore.PageMargins GetNativeMargins(WordSection section, PdfSaveOptions? options, (double Header, double Footer) headerFooterMarginExpansion) {
            if (options?.Margins != null) {
                return options.Margins.Value;
            }

            return new PdfCore.PageMargins(
                (section.Margins.Left?.Value ?? 0) / 20D,
                (section.Margins.Top ?? 0) / 20D + headerFooterMarginExpansion.Header,
                (section.Margins.Right?.Value ?? 0) / 20D,
                (section.Margins.Bottom ?? 0) / 20D + headerFooterMarginExpansion.Footer);
        }

        private static (double Header, double Footer) GetNativeHeaderFooterMarginExpansion(WordSection section, PdfSaveOptions? options) {
            if (options?.Margins != null) {
                return (0D, 0D);
            }

            double headerExpansion = GetNativeHeaderFooterMarginExpansion(
                section.Header?.Default,
                section.DifferentFirstPage ? section.Header?.First : null,
                section.DifferentOddAndEvenPages ? section.Header?.Even : null);
            double footerExpansion = GetNativeFooterMarginExpansion(
                section.Footer?.Default,
                section.DifferentFirstPage ? section.Footer?.First : null,
                section.DifferentOddAndEvenPages ? section.Footer?.Even : null);

            return (headerExpansion, footerExpansion);
        }

        private static double GetNativeHeaderFooterMarginExpansion(params WordHeaderFooter?[] variants) {
            int maxLines = 0;
            foreach (WordHeaderFooter? variant in variants) {
                maxLines = Math.Max(maxLines, GetNativeHeaderFooterLineCount(variant));
            }

            return GetNativeHeaderFooterMarginExpansion(maxLines, GetNativeHeaderFooterLineHeight(variants));
        }

        private static double GetNativeFooterMarginExpansion(params WordHeaderFooter?[] variants) {
            int maxLines = 0;
            foreach (WordHeaderFooter? variant in variants) {
                maxLines = Math.Max(maxLines, GetNativeHeaderFooterLineCount(variant));
            }

            return GetNativeFooterMarginExpansion(maxLines, GetNativeHeaderFooterLineHeight(variants));
        }

        private static double GetNativeHeaderFooterTextMarginExpansion(params NativeHeaderFooterText?[] variants) {
            int maxLines = 0;
            foreach (NativeHeaderFooterText? variant in variants) {
                maxLines = Math.Max(maxLines, GetNativeHeaderFooterTextLineCount(variant));
            }

            return GetNativeHeaderFooterMarginExpansion(maxLines);
        }

        private static double GetNativeHeaderFooterMarginExpansion(int maxLines, double lineHeight = NativeHeaderFooterLineHeight) {
            if (maxLines <= 2) {
                return 0D;
            }

            return (maxLines - 2) * lineHeight + NativeHeaderFooterBodyGap;
        }

        private static double GetNativeFooterMarginExpansion(int maxLines, double lineHeight = NativeHeaderFooterLineHeight) {
            if (maxLines <= 1) {
                return 0D;
            }

            return ((maxLines - 1) * lineHeight) + NativeHeaderFooterBodyGap;
        }

        private static double GetNativeHeaderFooterLineHeight(params WordHeaderFooter?[] variants) {
            double maxFontSize = NativeHeaderFooterFontSize;
            foreach (WordHeaderFooter? variant in variants) {
                foreach (double fontSize in EnumerateNativeHeaderFooterFontSizes(variant)) {
                    if (fontSize > maxFontSize) {
                        maxFontSize = fontSize;
                    }
                }
            }

            return maxFontSize * 1.2D;
        }

        private static int GetNativeHeaderFooterLineCount(WordHeaderFooter? headerFooter) {
            if (headerFooter == null) {
                return 0;
            }

            int textLines = GetNativeHeaderFooterTextLineCount(GetNativeHeaderFooterText(headerFooter));
            int structuralLines = 0;
            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                structuralLines += GetNativeHeaderFooterElementLineCount(element);
            }

            return Math.Max(textLines, structuralLines);
        }

        private static int GetNativeHeaderFooterElementLineCount(WordElement element) {
            return element switch {
                WordParagraph paragraph => GetNativeHeaderFooterParagraphLineCount(paragraph),
                WordTable table => GetNativeHeaderFooterTableLineCount(table),
                WordHyperLink link when !string.IsNullOrWhiteSpace(link.Text) => 1,
                _ => 0
            };
        }

        private static int GetNativeHeaderFooterTableLineCount(WordTable table) {
            int lineCount = 0;
            foreach (WordTableRow row in table.Rows) {
                int rowLineCount = 0;
                foreach (WordTableCell cell in row.Cells) {
                    int cellLineCount = 0;
                    foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                        cellLineCount += GetNativeHeaderFooterParagraphLineCount(paragraph);
                    }

                    rowLineCount = Math.Max(rowLineCount, cellLineCount);
                }

                lineCount += rowLineCount;
            }

            return lineCount;
        }

        private static int GetNativeHeaderFooterParagraphLineCount(WordParagraph paragraph) {
            string? text = GetNativeHeaderFooterParagraphText(paragraph, out _);
            return Math.Max(1, CountNativeHeaderFooterLines(text));
        }

        private static int GetNativeHeaderFooterTextLineCount(NativeHeaderFooterText? text) {
            if (text == null) {
                return 0;
            }

            return Math.Max(
                CountNativeHeaderFooterLines(text.Left),
                Math.Max(
                    CountNativeHeaderFooterLines(text.Center),
                    CountNativeHeaderFooterLines(text.Right)));
        }

        private static int CountNativeHeaderFooterLines(string? text) {
            if (string.IsNullOrEmpty(text)) {
                return 0;
            }

            string normalized = text!.Replace("\r\n", "\n").Replace('\r', '\n');
            int lines = 1;
            for (int index = 0; index < normalized.Length; index++) {
                if (normalized[index] == '\n') {
                    lines++;
                }
            }

            return lines;
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
            int metadataSeparator = value.IndexOf(' ');
            if (metadataSeparator > 0) {
                value = value.Substring(0, metadataSeparator);
            }

            if (value.Equals("none", StringComparison.OrdinalIgnoreCase) ||
                !OfficeColor.TryParse(value, out OfficeColor color)) {
                return null;
            }

            return PdfCore.PdfColor.FromOfficeColor(color);
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
