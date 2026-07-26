using AngleSharp.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
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
            internal BlockBorder(BorderValues style, UInt32Value size, SixColor? color) {
                Style = style;
                Size = size;
                Color = color;
            }

            internal BorderValues Style { get; }
            internal UInt32Value Size { get; }
            internal SixColor? Color { get; }
        }

        private struct BlockBorderState {
            internal BlockBorderState(
                BorderValues style,
                UInt32Value size,
                SixColor? color,
                bool transparentColor = false) {
                Style = style;
                Size = size;
                Color = color;
                TransparentColor = transparentColor;
            }

            internal BorderValues? Style { get; set; }
            internal UInt32Value? Size { get; set; }
            internal SixColor? Color { get; set; }
            internal bool TransparentColor { get; set; }

            internal BlockBorder? Materialize() =>
                !TransparentColor &&
                Style.HasValue &&
                Style.Value != BorderValues.None &&
                (Size == null || Size.Value > 0)
                    ? new BlockBorder(Style.Value, Size ?? 4U, Color)
                    : null;
        }

        private static void ApplyParagraphFrameFromCss(WordParagraph paragraph, IElement element) {
            ApplyBlockFrameFromCss(element, new[] { paragraph }, applyContainerSpacing: false);
        }

        private static CssStyleMapper.CssProperties ParseElementBoxStyles(IElement element) {
            var lineage = new Stack<IElement>();
            for (IElement? current = element; current != null; current = current.ParentElement) {
                lineage.Push(current);
            }

            CssStyleMapper.CssProperties? inheritedBox = null;
            bool inheritedRightToLeft = false;
            while (lineage.Count > 0) {
                IElement current = lineage.Pop();
                bool rightToLeft = GetBidiFromDir(current) == true;
                inheritedBox = CssStyleMapper.ParseStyles(
                    current.GetAttribute("style"),
                    rightToLeft,
                    inheritedBox,
                    inheritedRightToLeft);
                inheritedRightToLeft = rightToLeft;
            }

            return inheritedBox ?? new CssStyleMapper.CssProperties();
        }

        private static bool RequiresEmptyBlockParagraph(IElement element) {
            if (StyleRequestsPageBreakBefore(element) || StyleRequestsPageBreakAfter(element)) {
                return true;
            }

            CssStyleMapper.CssProperties box = ParseElementBoxStyles(element);
            if ((box.MarginTop ?? 0) != 0 ||
                (box.MarginBottom ?? 0) != 0 ||
                (box.PaddingTop ?? 0) != 0 ||
                (box.PaddingRight ?? 0) != 0 ||
                (box.PaddingBottom ?? 0) != 0 ||
                (box.PaddingLeft ?? 0) != 0) {
                return true;
            }

            if (string.IsNullOrWhiteSpace(element.GetAttribute("style"))) {
                return false;
            }

            ParseElementBlockFrameStyles(element, out string? background, out var sideBorders);
            return background != null || sideBorders.Values.Any(border => border.Materialize().HasValue);
        }

        private static void ParseElementBlockFrameStyles(
            IElement element,
            out string? background,
            out Dictionary<BlockBorderSide, BlockBorderState> sideBorders) {
            var lineage = new Stack<IElement>();
            for (IElement? current = element; current != null; current = current.ParentElement) {
                lineage.Push(current);
            }

            string? inheritedBackground = null;
            var inheritedBorders = new Dictionary<BlockBorderSide, BlockBorderState>();
            while (lineage.Count > 0) {
                IElement current = lineage.Pop();
                ParseBlockFrameStyles(
                    current.GetAttribute("style") ?? string.Empty,
                    inheritedBackground,
                    inheritedBorders,
                    out background,
                    out sideBorders);
                inheritedBackground = background;
                inheritedBorders = sideBorders;
            }

            background = inheritedBackground;
            sideBorders = inheritedBorders;
        }

        private static void ParseBlockFrameStyles(
            string styleText,
            string? inheritedBackground,
            IReadOnlyDictionary<BlockBorderSide, BlockBorderState> inheritedBorders,
            out string? background,
            out Dictionary<BlockBorderSide, BlockBorderState> sideBorders) {
            background = ResolveBlockBackground(styleText, inheritedBackground);
            sideBorders = new Dictionary<BlockBorderSide, BlockBorderState>();
            for (int priorityPass = 0; priorityPass < 2; priorityPass++) {
                bool important = priorityPass == 1;
                foreach (string part in styleText.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                    if (!CssStyleMapper.TryParseDeclaration(
                            part,
                            out string name,
                            out string value,
                            out bool declarationIsImportant) ||
                        declarationIsImportant != important) {
                        continue;
                    }

                    if (name == "background-color") {
                        continue;
                    } else if (name == "border" && IsCssInheritanceValue(value)) {
                        CopyAllBlockBorders(inheritedBorders, sideBorders);
                    } else if (name == "border" && IsCssWideResetValue(value)) {
                        ResetAllBlockBorders(sideBorders);
                    } else if (name == "border" &&
                               TryParseBlockBorder(
                                   value,
                                   out var borderStyle,
                                   out var borderSize,
                                   out var borderColor,
                                   out bool hasBorderStyle,
                                   out bool hasBorderColor,
                                   out bool hasTransparentBorderColor,
                                   background ?? inheritedBackground)) {
                        var border = new BlockBorderState(
                            hasBorderStyle ? borderStyle : BorderValues.None,
                            borderSize,
                            hasBorderColor ? borderColor : null,
                            hasTransparentBorderColor);
                        sideBorders[BlockBorderSide.Left] = border;
                        sideBorders[BlockBorderSide.Right] = border;
                        sideBorders[BlockBorderSide.Top] = border;
                        sideBorders[BlockBorderSide.Bottom] = border;
                    } else if (TryGetBlockBorderSide(name, out BlockBorderSide side) &&
                               IsCssInheritanceValue(value)) {
                        CopyBlockBorder(inheritedBorders, sideBorders, side);
                    } else if (TryGetBlockBorderSide(name, out side) &&
                               IsCssWideResetValue(value)) {
                        sideBorders[side] = default;
                    } else if (TryGetBlockBorderSide(name, out side) &&
                               TryParseBlockBorder(
                                   value,
                                   out var sideStyle,
                                   out var sideSize,
                                   out var sideColor,
                                   out bool hasSideStyle,
                                   out bool hasSideColor,
                                   out bool hasTransparentSideColor,
                                   background ?? inheritedBackground)) {
                        sideBorders[side] = new BlockBorderState(
                            hasSideStyle ? sideStyle : BorderValues.None,
                            sideSize,
                            hasSideColor ? sideColor : null,
                            hasTransparentSideColor);
                    } else if (TryGetBlockBorderLonghand(name, out side, out string component)) {
                        ApplyBlockBorderLonghand(
                            sideBorders,
                            inheritedBorders,
                            side,
                            component,
                            value,
                            background ?? inheritedBackground);
                    }
                }
            }
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

        private static void ApplyContainerFrameFromCss(
            IElement element,
            IReadOnlyList<WordParagraph> paragraphs,
            IReadOnlyList<WordTable> tables) {
            ApplyContainerFrameFromCss(element, paragraphs, applyContainerSpacing: false);
            foreach (WordTable table in tables) {
                ApplyTableContainerFrameFromCss(element, table);
            }
            ApplyContainerSpacingFromCss(element, paragraphs, tables);
        }

        private static void ApplyTableContainerFrameFromCss(IElement element, WordTable table) {
            List<WordTableRow> rows = table.Rows;
            if (rows.Count == 0) {
                return;
            }

            ParseElementBlockFrameStyles(element, out string? background, out var sideBorders);
            foreach (WordTableRow row in rows) {
                foreach (WordTableCell cell in row.Cells) {
                    if (!string.IsNullOrEmpty(background) && string.IsNullOrEmpty(cell.ShadingFillColorHex)) {
                        cell.ShadingFillColorHex = background!;
                    }
                }
            }

            BlockBorder? left = GetBlockBorder(sideBorders, BlockBorderSide.Left);
            BlockBorder? right = GetBlockBorder(sideBorders, BlockBorderSide.Right);
            BlockBorder? top = GetBlockBorder(sideBorders, BlockBorderSide.Top);
            BlockBorder? bottom = GetBlockBorder(sideBorders, BlockBorderSide.Bottom);
            string fallbackColor = ResolveElementTextColor(element);

            foreach (WordTableCell cell in rows[0].Cells) {
                ApplyTableCellBorder(cell, BlockBorderSide.Top, top, fallbackColor);
            }
            foreach (WordTableCell cell in rows[rows.Count - 1].Cells) {
                ApplyTableCellBorder(cell, BlockBorderSide.Bottom, bottom, fallbackColor);
            }
            foreach (WordTableRow row in rows) {
                List<WordTableCell> cells = row.Cells;
                if (cells.Count == 0) {
                    continue;
                }
                ApplyTableCellBorder(cells[0], BlockBorderSide.Left, left, fallbackColor);
                ApplyTableCellBorder(cells[cells.Count - 1], BlockBorderSide.Right, right, fallbackColor);
            }
        }

        private static void ApplyTableCellBorder(
            WordTableCell cell,
            BlockBorderSide side,
            BlockBorder? border,
            string fallbackColor) {
            if (!border.HasValue || HasTableCellBorder(cell, side)) {
                return;
            }

            BlockBorder value = border.Value;
            string color = value.Color?.ToRgbHex() ?? fallbackColor;
            switch (side) {
                case BlockBorderSide.Left:
                    cell.Borders.LeftStyle = value.Style;
                    cell.Borders.LeftSize = value.Size;
                    cell.Borders.LeftColorHex = color;
                    break;
                case BlockBorderSide.Right:
                    cell.Borders.RightStyle = value.Style;
                    cell.Borders.RightSize = value.Size;
                    cell.Borders.RightColorHex = color;
                    break;
                case BlockBorderSide.Top:
                    cell.Borders.TopStyle = value.Style;
                    cell.Borders.TopSize = value.Size;
                    cell.Borders.TopColorHex = color;
                    break;
                case BlockBorderSide.Bottom:
                    cell.Borders.BottomStyle = value.Style;
                    cell.Borders.BottomSize = value.Size;
                    cell.Borders.BottomColorHex = color;
                    break;
            }
        }

        private static bool HasTableCellBorder(WordTableCell cell, BlockBorderSide side) =>
            side switch {
                BlockBorderSide.Left => cell.Borders.LeftStyle.HasValue,
                BlockBorderSide.Right => cell.Borders.RightStyle.HasValue,
                BlockBorderSide.Top => cell.Borders.TopStyle.HasValue,
                BlockBorderSide.Bottom => cell.Borders.BottomStyle.HasValue,
                _ => false,
            };

        private static string ResolveElementTextColor(IElement element) {
            string styleText = element.GetAttribute("style") ?? string.Empty;
            var declaration = ParseInlineDeclaration(styleText);
            return NormalizeColor(GetInlinePropertyValue(declaration, styleText, "color")) ??
                   SixColor.Black.ToRgbHex();
        }

        private static void ApplyBlockFrameFromCss(
            IElement element,
            IReadOnlyList<WordParagraph> paragraphs,
            bool applyContainerSpacing) {
            string? styleText = element.GetAttribute("style");
            if (string.IsNullOrWhiteSpace(styleText) || paragraphs.Count == 0) {
                return;
            }

            paragraphs = DistinctPhysicalParagraphs(paragraphs);
            ParseElementBlockFrameStyles(element, out string? background, out var sideBorders);

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
            string fallbackColor = ResolveElementTextColor(element);
            for (int index = 0; index < paragraphs.Count; index++) {
                WordParagraph paragraph = paragraphs[index];
                ApplyParagraphBorder(paragraph, BlockBorderSide.Left, left, fallbackColor);
                ApplyParagraphBorder(paragraph, BlockBorderSide.Right, right, fallbackColor);
                if (index == 0) {
                    ApplyParagraphBorder(paragraph, BlockBorderSide.Top, top, fallbackColor);
                }
                if (index == paragraphs.Count - 1) {
                    ApplyParagraphBorder(paragraph, BlockBorderSide.Bottom, bottom, fallbackColor);
                }
            }

            if (!applyContainerSpacing) {
                return;
            }

            CssStyleMapper.CssProperties box = ParseElementBoxStyles(element);
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

        private static IReadOnlyList<WordParagraph> DistinctPhysicalParagraphs(
            IReadOnlyList<WordParagraph> paragraphs) {
            var result = new List<WordParagraph>(paragraphs.Count);
            foreach (WordParagraph paragraph in paragraphs) {
                if (!result.Any(existing => ReferenceEquals(existing._paragraph, paragraph._paragraph))) {
                    result.Add(paragraph);
                }
            }
            return result;
        }

        private static BlockBorder? GetBlockBorder(
            IReadOnlyDictionary<BlockBorderSide, BlockBorderState> sideBorders,
            BlockBorderSide side) =>
            sideBorders.TryGetValue(side, out BlockBorderState border) ? border.Materialize() : null;

        private static void ApplyBlockBorderLonghand(
            IDictionary<BlockBorderSide, BlockBorderState> sideBorders,
            IReadOnlyDictionary<BlockBorderSide, BlockBorderState> inheritedBorders,
            BlockBorderSide side,
            string component,
            string value,
            string? backdrop) {
            BlockBorderState border = sideBorders.TryGetValue(side, out BlockBorderState existing)
                ? existing
                : default;
            if (IsCssInheritanceValue(value)) {
                BlockBorderState inherited = inheritedBorders.TryGetValue(side, out BlockBorderState inheritedBorder)
                    ? inheritedBorder
                    : default;
                switch (component) {
                    case "style":
                        border.Style = inherited.Style;
                        break;
                    case "width":
                        border.Size = inherited.Size;
                        break;
                    case "color":
                        border.Color = inherited.Color;
                        border.TransparentColor = inherited.TransparentColor;
                        break;
                    default:
                        return;
                }
                sideBorders[side] = border;
                return;
            }
            if (IsCssWideResetValue(value)) {
                switch (component) {
                    case "style":
                        border.Style = BorderValues.None;
                        break;
                    case "width":
                        border.Size = null;
                        break;
                    case "color":
                        border.Color = null;
                        border.TransparentColor = false;
                        break;
                    default:
                        return;
                }
                sideBorders[side] = border;
                return;
            }
            switch (component) {
                case "style":
                    if (!TryParseBlockBorderStyle(value, out BorderValues style)) {
                        return;
                    }
                    border.Style = style;
                    break;
                case "width":
                    if (!TryParseBlockBorderWidth(value, out UInt32Value size)) {
                        return;
                    }
                    border.Size = size;
                    break;
                case "color":
                    if (!TryResolveBlockBorderColor(
                            value,
                            backdrop,
                            out SixColor color,
                            out bool transparentColor)) {
                        return;
                    }
                    border.Color = transparentColor ? null : color;
                    border.TransparentColor = transparentColor;
                    break;
                default:
                    return;
            }
            sideBorders[side] = border;
        }

        private static bool IsCssInheritanceValue(string value) =>
            value.Trim().Equals("inherit", StringComparison.OrdinalIgnoreCase);

        private static bool IsCssWideResetValue(string value) {
            string normalized = value.Trim().ToLowerInvariant();
            return normalized is "initial" or "unset" or "revert" or "revert-layer";
        }

        private static void CopyAllBlockBorders(
            IReadOnlyDictionary<BlockBorderSide, BlockBorderState> source,
            IDictionary<BlockBorderSide, BlockBorderState> destination) {
            foreach (BlockBorderSide side in Enum.GetValues(typeof(BlockBorderSide))) {
                CopyBlockBorder(source, destination, side);
            }
        }

        private static void CopyBlockBorder(
            IReadOnlyDictionary<BlockBorderSide, BlockBorderState> source,
            IDictionary<BlockBorderSide, BlockBorderState> destination,
            BlockBorderSide side) {
            destination[side] = source.TryGetValue(side, out BlockBorderState inherited)
                ? inherited
                : default;
        }

        private static void ResetAllBlockBorders(IDictionary<BlockBorderSide, BlockBorderState> sideBorders) {
            foreach (BlockBorderSide side in Enum.GetValues(typeof(BlockBorderSide))) {
                sideBorders[side] = default;
            }
        }

        private static bool TryParseBlockBorder(
            string value,
            out BorderValues style,
            out UInt32Value size,
            out SixColor color,
            out bool hasExplicitStyle,
            out bool hasExplicitColor,
            out bool hasTransparentColor,
            string? backdrop) {
            hasTransparentColor = false;
            string[] tokens = SplitBorderTokens(value).ToArray();
            string normalized = string.Join(
                " ",
                tokens.Select(token => token.Trim() == "0" ? "0px" : token));
            if (!TryParseBorder(
                    normalized,
                    out style,
                    out size,
                    out color,
                    out hasExplicitStyle,
                    out hasExplicitColor)) {
                return false;
            }

            foreach (string token in tokens) {
                if (!TryParseRawBlockBorderWidth(token, out double width)) {
                    if (hasExplicitColor &&
                        !TryParseBlockBorderStyle(token, out _) &&
                        TryResolveBlockBorderColor(
                            token,
                            backdrop,
                            out SixColor resolvedColor,
                            out bool transparentColor)) {
                        color = resolvedColor;
                        hasTransparentColor = transparentColor;
                    }
                } else {
                    if (width < 0) {
                        return false;
                    }
                    if (width == 0) {
                        size = 0U;
                    }
                }
            }

            return true;
        }

        private static bool TryParseBlockBorderWidth(string value, out UInt32Value size) {
            if (!TryParseRawBlockBorderWidth(value, out double width) || width < 0) {
                size = 0U;
                return false;
            }
            if (width == 0) {
                size = 0U;
                return true;
            }

            return TryParseBorderWidth(value, out size);
        }

        private static bool TryParseRawBlockBorderWidth(string value, out double width) {
            string normalized = value.Trim().ToLowerInvariant();
            if (normalized == "0") {
                width = 0;
                return true;
            }

            if (normalized.EndsWith("px", StringComparison.Ordinal) ||
                normalized.EndsWith("pt", StringComparison.Ordinal)) {
                return double.TryParse(
                    normalized.Substring(0, normalized.Length - 2),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out width);
            }

            width = 0;
            return false;
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
            BlockBorder? border,
            string fallbackColor) {
            if (!border.HasValue || HasParagraphBorder(paragraph, side)) {
                return;
            }

            BlockBorder value = border.Value;
            string color = value.Color?.ToRgbHex() ?? fallbackColor;
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

        private static bool HasParagraphBorder(WordParagraph paragraph, BlockBorderSide side) =>
            side switch {
                BlockBorderSide.Left => paragraph.Borders.LeftStyle.HasValue,
                BlockBorderSide.Right => paragraph.Borders.RightStyle.HasValue,
                BlockBorderSide.Top => paragraph.Borders.TopStyle.HasValue,
                BlockBorderSide.Bottom => paragraph.Borders.BottomStyle.HasValue,
                _ => false,
            };

    }
}
