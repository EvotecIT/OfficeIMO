using AngleSharp.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using SixColor = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private enum BlockBorderSide {
            Left,
            Right,
            Top,
            Bottom
        }

        private readonly struct BlockBorder {
            internal BlockBorder(BorderValues style, UInt32Value size, SixColor color) {
                Style = style;
                Size = size;
                Color = color;
            }

            internal BorderValues Style { get; }
            internal UInt32Value Size { get; }
            internal SixColor Color { get; }
        }

        private struct BlockBorderState {
            internal BlockBorderState(BorderValues style, UInt32Value size, SixColor color) {
                Style = style;
                Size = size;
                Color = color;
            }

            internal BorderValues? Style { get; set; }
            internal UInt32Value? Size { get; set; }
            internal SixColor? Color { get; set; }

            internal BlockBorder? Materialize() =>
                Style.HasValue
                    ? new BlockBorder(Style.Value, Size ?? 4U, Color ?? SixColor.Black)
                    : null;
        }

        private static void ApplyParagraphFrameFromCss(WordParagraph paragraph, IElement element) {
            ApplyBlockFrameFromCss(element, new[] { paragraph }, applyContainerSpacing: false);
        }

        private static void ApplyContainerFrameFromCss(
            IElement element,
            IReadOnlyList<WordParagraph> paragraphs,
            bool applyContainerSpacing = true) {
            if (paragraphs.Count == 0) {
                return;
            }

            ApplyBlockFrameFromCss(element, paragraphs, applyContainerSpacing);
        }

        private static void ApplyBlockFrameFromCss(
            IElement element,
            IReadOnlyList<WordParagraph> paragraphs,
            bool applyContainerSpacing) {
            string? styleText = element.GetAttribute("style");
            if (string.IsNullOrWhiteSpace(styleText) || paragraphs.Count == 0) {
                return;
            }

            string? background = null;
            var sideBorders = new Dictionary<BlockBorderSide, BlockBorderState>();
            for (int priorityPass = 0; priorityPass < 2; priorityPass++) {
                bool important = priorityPass == 1;
                foreach (string part in styleText!.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                    if (!CssStyleMapper.TryParseDeclaration(
                            part,
                            out string name,
                            out string value,
                            out bool declarationIsImportant) ||
                        declarationIsImportant != important) {
                        continue;
                    }

                    if (name == "background-color") {
                        if (value.Equals("transparent", StringComparison.OrdinalIgnoreCase)) {
                            background = null;
                        } else {
                            string? normalizedBackground = NormalizeColor(value);
                            if (normalizedBackground != null) {
                                background = normalizedBackground;
                            }
                        }
                    } else if (name == "border" &&
                               TryParseBorder(value, out var borderStyle, out var borderSize, out var borderColor, out bool hasBorderStyle)) {
                        if (hasBorderStyle) {
                            var border = new BlockBorderState(borderStyle, borderSize, borderColor);
                            sideBorders[BlockBorderSide.Left] = border;
                            sideBorders[BlockBorderSide.Right] = border;
                            sideBorders[BlockBorderSide.Top] = border;
                            sideBorders[BlockBorderSide.Bottom] = border;
                        } else {
                            sideBorders.Clear();
                        }
                    } else if (TryGetBlockBorderSide(name, out BlockBorderSide side) &&
                               TryParseBorder(value, out var sideStyle, out var sideSize, out var sideColor, out bool hasSideStyle)) {
                        if (hasSideStyle) {
                            sideBorders[side] = new BlockBorderState(sideStyle, sideSize, sideColor);
                        } else {
                            sideBorders.Remove(side);
                        }
                    } else if (TryGetBlockBorderLonghand(name, out side, out string component)) {
                        ApplyBlockBorderLonghand(sideBorders, side, component, value);
                    }
                }
            }

            if (!string.IsNullOrEmpty(background)) {
                foreach (WordParagraph paragraph in paragraphs) {
                    if (string.IsNullOrEmpty(paragraph.ShadingFillColorHex)) {
                        paragraph.ShadingFillColorHex = background!;
                    }
                }
            }

            BlockBorder? left = GetBlockBorder(sideBorders, BlockBorderSide.Left);
            BlockBorder? right = GetBlockBorder(sideBorders, BlockBorderSide.Right);
            BlockBorder? top = GetBlockBorder(sideBorders, BlockBorderSide.Top);
            BlockBorder? bottom = GetBlockBorder(sideBorders, BlockBorderSide.Bottom);
            for (int index = 0; index < paragraphs.Count; index++) {
                WordParagraph paragraph = paragraphs[index];
                ApplyParagraphBorder(paragraph, BlockBorderSide.Left, left);
                ApplyParagraphBorder(paragraph, BlockBorderSide.Right, right);
                if (index == 0) {
                    ApplyParagraphBorder(paragraph, BlockBorderSide.Top, top);
                }
                if (index == paragraphs.Count - 1) {
                    ApplyParagraphBorder(paragraph, BlockBorderSide.Bottom, bottom);
                }
            }

            if (!applyContainerSpacing) {
                return;
            }

            CssStyleMapper.CssProperties box = CssStyleMapper.ParseStyles(styleText, GetBidiFromDir(element) == true);
            int horizontalStart = (box.MarginLeft ?? 0) + (box.PaddingLeft ?? 0);
            int horizontalEnd = (box.MarginRight ?? 0) + (box.PaddingRight ?? 0);
            foreach (WordParagraph paragraph in paragraphs) {
                if (horizontalStart != 0) {
                    paragraph.IndentationBefore = (paragraph.IndentationBefore ?? 0) + horizontalStart;
                }
                if (horizontalEnd != 0) {
                    paragraph.IndentationAfter = (paragraph.IndentationAfter ?? 0) + horizontalEnd;
                }
            }

            int verticalStart = (box.MarginTop ?? 0) + (box.PaddingTop ?? 0);
            if (verticalStart != 0) {
                paragraphs[0].LineSpacingBefore = (paragraphs[0].LineSpacingBefore ?? 0) + verticalStart;
            }
            int verticalEnd = (box.MarginBottom ?? 0) + (box.PaddingBottom ?? 0);
            if (verticalEnd != 0) {
                WordParagraph last = paragraphs[paragraphs.Count - 1];
                last.LineSpacingAfter = (last.LineSpacingAfter ?? 0) + verticalEnd;
            }
        }

        private static BlockBorder? GetBlockBorder(
            IReadOnlyDictionary<BlockBorderSide, BlockBorderState> sideBorders,
            BlockBorderSide side) =>
            sideBorders.TryGetValue(side, out BlockBorderState border) ? border.Materialize() : null;

        private static void ApplyBlockBorderLonghand(
            IDictionary<BlockBorderSide, BlockBorderState> sideBorders,
            BlockBorderSide side,
            string component,
            string value) {
            BlockBorderState border = sideBorders.TryGetValue(side, out BlockBorderState existing)
                ? existing
                : default;
            switch (component) {
                case "style":
                    if (!TryParseBlockBorderStyle(value, out BorderValues style)) {
                        return;
                    }
                    border.Style = style;
                    break;
                case "width":
                    if (!TryParseBorderWidth(value, out UInt32Value size)) {
                        return;
                    }
                    border.Size = size;
                    break;
                case "color":
                    string? color = NormalizeColor(value);
                    if (color == null) {
                        return;
                    }
                    border.Color = SixColor.Parse("#" + color);
                    break;
                default:
                    return;
            }
            sideBorders[side] = border;
        }

        private static bool TryParseBlockBorderStyle(string value, out BorderValues style) {
            switch (value.Trim().ToLowerInvariant()) {
                case "solid":
                    style = BorderValues.Single;
                    return true;
                case "dotted":
                    style = BorderValues.Dotted;
                    return true;
                case "dashed":
                    style = BorderValues.Dashed;
                    return true;
                case "double":
                    style = BorderValues.Double;
                    return true;
                case "none":
                    style = BorderValues.None;
                    return true;
                default:
                    style = BorderValues.Single;
                    return false;
            }
        }

        private static bool TryGetBlockBorderLonghand(
            string propertyName,
            out BlockBorderSide side,
            out string component) {
            string[] pieces = propertyName.Split('-');
            if (pieces.Length == 3 &&
                pieces[0].Equals("border", StringComparison.OrdinalIgnoreCase) &&
                TryParseBlockBorderSide(pieces[1], out side) &&
                (pieces[2].Equals("style", StringComparison.OrdinalIgnoreCase) ||
                 pieces[2].Equals("width", StringComparison.OrdinalIgnoreCase) ||
                 pieces[2].Equals("color", StringComparison.OrdinalIgnoreCase))) {
                component = pieces[2].ToLowerInvariant();
                return true;
            }

            side = default;
            component = string.Empty;
            return false;
        }

        private static bool TryParseBlockBorderSide(string value, out BlockBorderSide side) {
            switch (value.ToLowerInvariant()) {
                case "left":
                    side = BlockBorderSide.Left;
                    return true;
                case "right":
                    side = BlockBorderSide.Right;
                    return true;
                case "top":
                    side = BlockBorderSide.Top;
                    return true;
                case "bottom":
                    side = BlockBorderSide.Bottom;
                    return true;
                default:
                    side = default;
                    return false;
            }
        }

        private static bool TryGetBlockBorderSide(string propertyName, out BlockBorderSide side) {
            switch (propertyName) {
                case "border-left":
                    side = BlockBorderSide.Left;
                    return true;
                case "border-right":
                    side = BlockBorderSide.Right;
                    return true;
                case "border-top":
                    side = BlockBorderSide.Top;
                    return true;
                case "border-bottom":
                    side = BlockBorderSide.Bottom;
                    return true;
                default:
                    side = default;
                    return false;
            }
        }

        private static void ApplyParagraphBorder(
            WordParagraph paragraph,
            BlockBorderSide side,
            BlockBorder? border) {
            if (!border.HasValue) {
                return;
            }

            BlockBorder value = border.Value;
            string color = value.Color.ToRgbHex();
            switch (side) {
                case BlockBorderSide.Left:
                    paragraph.Borders.LeftStyle = value.Style;
                    paragraph.Borders.LeftSize = value.Size;
                    paragraph.Borders.LeftColorHex = color;
                    break;
                case BlockBorderSide.Right:
                    paragraph.Borders.RightStyle = value.Style;
                    paragraph.Borders.RightSize = value.Size;
                    paragraph.Borders.RightColorHex = color;
                    break;
                case BlockBorderSide.Top:
                    paragraph.Borders.TopStyle = value.Style;
                    paragraph.Borders.TopSize = value.Size;
                    paragraph.Borders.TopColorHex = color;
                    break;
                case BlockBorderSide.Bottom:
                    paragraph.Borders.BottomStyle = value.Style;
                    paragraph.Borders.BottomSize = value.Size;
                    paragraph.Borders.BottomColorHex = color;
                    break;
            }
        }
    }
}
