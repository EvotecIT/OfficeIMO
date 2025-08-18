using AngleSharp.Html.Dom;
using AngleSharp.Dom;
using AngleSharp.Css.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using SixLabors.ImageSharp;
using SixColor = SixLabors.ImageSharp.Color;
using System;
using System.Collections.Generic;

namespace OfficeIMO.Word.Html.Converters {
    internal partial class HtmlToWordConverter {
        private void ProcessTable(IHtmlTableElement tableElem, WordDocument doc, WordSection section, HtmlToWordOptions options,
            Stack<WordList> listStack, WordTableCell? cell, WordParagraph? currentParagraph, WordHeaderFooter? headerFooter) {
            int headRows = tableElem.Head?.Rows.Length ?? 0;
            int bodyRows = 0;
            foreach (var body in tableElem.Bodies) {
                bodyRows += body.Rows.Length;
            }
            int footRows = tableElem.Foot?.Rows.Length ?? 0;
            int rows = headRows + bodyRows + footRows;

            var caption = tableElem.Caption;
            if (caption != null) {
                ApplyCssToElement(caption);
            }

            int cols = 0;
            foreach (var row in GetAllRows(tableElem)) {
                int count = 0;
                foreach (var cellElem in row.Cells) {
                    int span = 1;
                    if (cellElem is IHtmlTableCellElement cellElement) {
                        span = Math.Max(1, cellElement.ColumnSpan);
                    }
                    count += span;
                }
                cols = Math.Max(cols, count);
            }
            WordParagraph? captionParagraph = null;
            if (caption != null && options.TableCaptionPosition == TableCaptionPosition.Above) {
                captionParagraph = cell is not null ? cell.AddParagraph("", true)
                    : currentParagraph is not null ? currentParagraph.AddParagraphAfterSelf()
                    : headerFooter is not null ? headerFooter.AddParagraph("")
                    : section.AddParagraph("");
                captionParagraph.SetStyleId("Caption");
                var props = ApplyParagraphStyleFromCss(captionParagraph, caption);
                ApplyClassStyle(caption, captionParagraph, options);
                AddBookmarkIfPresent(caption, captionParagraph);
                var fmt = new TextFormatting();
                if (props.WhiteSpace.HasValue) {
                    fmt.WhiteSpace = props.WhiteSpace.Value;
                }
                foreach (var child in caption.ChildNodes) {
                    ProcessNode(child, doc, section, options, captionParagraph, listStack, fmt, cell, headerFooter);
                }
            }

            WordTable wordTable;
            if (captionParagraph is not null) {
                wordTable = captionParagraph.AddTableAfter(rows, cols);
            } else if (cell is not null) {
                wordTable = cell.AddTable(rows, cols);
            } else if (currentParagraph is not null) {
                wordTable = currentParagraph.AddTableAfter(rows, cols);
            } else {
                var placeholder = headerFooter is not null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                wordTable = placeholder.AddTableAfter(rows, cols);
            }
            ApplyTableStyles(wordTable, tableElem);
            ApplyColumnGroup(wordTable, tableElem, cols);
            var occupied = new bool[rows, cols];
            int rIndex = 0;

            void HandleRows(IEnumerable<IHtmlTableRowElement> htmlRows) {
                foreach (var htmlRow in htmlRows) {
                    var wordRow = wordTable.Rows[rIndex];
                    ApplyRowStyles(wordRow, htmlRow);
                    int cIndex = 0;
                    for (int c = 0; c < htmlRow.Cells.Length; c++) {
                        while (cIndex < cols && occupied[rIndex, cIndex]) {
                            cIndex++;
                        }

                        var htmlCell = htmlRow.Cells[c];
                        var wordCell = wordRow.Cells[cIndex];
                        var alignment = ApplyCellStyles(wordCell, htmlCell as IHtmlTableCellElement);
                        if (wordCell.Paragraphs.Count == 1 && string.IsNullOrEmpty(wordCell.Paragraphs[0].Text)) {
                            wordCell.Paragraphs[0].Remove();
                        }

                        WordParagraph? innerParagraph = null;
                        foreach (var child in htmlCell.ChildNodes) {
                            ProcessNode(child, doc, section, options, innerParagraph, listStack, new TextFormatting(), wordCell, headerFooter);
                            if (wordCell.Paragraphs.Count > 0) {
                                innerParagraph = wordCell.Paragraphs[wordCell.Paragraphs.Count - 1];
                            } else {
                                innerParagraph = null;
                            }
                        }

                        if (alignment.HasValue) {
                            foreach (var p in wordCell.Paragraphs) {
                                p.ParagraphAlignment = alignment.Value;
                            }
                        }

                        int rowSpan = 1;
                        int colSpan = 1;
                        if (htmlCell is IHtmlTableCellElement htmlTableCell) {
                            rowSpan = Math.Max(1, htmlTableCell.RowSpan);
                            colSpan = Math.Max(1, htmlTableCell.ColumnSpan);
                        }

                        if (rowSpan > 1 || colSpan > 1) {
                            wordTable.MergeCells(rIndex, cIndex, rowSpan, colSpan);
                            for (int rr = rIndex; rr < rIndex + rowSpan; rr++) {
                                for (int cc = cIndex; cc < cIndex + colSpan; cc++) {
                                    if (rr == rIndex && cc == cIndex) {
                                        continue;
                                    }
                                    occupied[rr, cc] = true;
                                }
                            }
                        }

                        cIndex++;
                    }
                    rIndex++;
                }
            }

            if (tableElem.Head != null) {
                HandleRows(tableElem.Head.Rows);
                if (tableElem.Head.Rows.Length > 0) {
                    wordTable.RepeatHeaderRowAtTheTopOfEachPage = true;
                }
            }
            foreach (var body in tableElem.Bodies) {
                HandleRows(body.Rows);
            }
            if (tableElem.Foot != null) {
                HandleRows(tableElem.Foot.Rows);
            }

            if (caption != null && options.TableCaptionPosition == TableCaptionPosition.Below) {
                WordParagraph captionParagraphBelow;
                if (cell is not null) {
                    captionParagraphBelow = cell.AddParagraph("", true);
                } else if (headerFooter is not null) {
                    captionParagraphBelow = headerFooter.AddParagraph("");
                } else {
                    var lastCellParagraph = wordTable.Rows[wordTable.Rows.Count - 1]
                        .Cells[wordTable.Rows[0].Cells.Count - 1].Paragraphs.Last();
                    captionParagraphBelow = lastCellParagraph.AddParagraphAfterSelf(section);
                }
                captionParagraphBelow.SetStyleId("Caption");
                var propsBelow = ApplyParagraphStyleFromCss(captionParagraphBelow, caption);
                ApplyClassStyle(caption, captionParagraphBelow, options);
                AddBookmarkIfPresent(caption, captionParagraphBelow);
                var fmtBelow = new TextFormatting();
                if (propsBelow.WhiteSpace.HasValue) {
                    fmtBelow.WhiteSpace = propsBelow.WhiteSpace.Value;
                }
                foreach (var child in caption.ChildNodes) {
                    ProcessNode(child, doc, section, options, captionParagraphBelow, listStack, fmtBelow, cell, headerFooter);
                }
            }
        }

        private static IEnumerable<IHtmlTableRowElement> GetAllRows(IHtmlTableElement tableElem) {
            if (tableElem.Head != null) {
                foreach (var r in tableElem.Head.Rows) yield return r;
            }
            foreach (var body in tableElem.Bodies) {
                foreach (var r in body.Rows) yield return r;
            }
            if (tableElem.Foot != null) {
                foreach (var r in tableElem.Foot.Rows) yield return r;
            }
        }

        private static void ApplyColumnGroup(WordTable wordTable, IHtmlTableElement tableElem, int cols) {
            var colElements = tableElem.QuerySelectorAll("col");
            if (colElements.Length == 0) {
                return;
            }
            var widths = new List<int>();
            TableWidthUnitValues? widthType = null;
            foreach (var col in colElements) {
                string? width = null;
                var style = col.GetAttribute("style");
                if (!string.IsNullOrEmpty(style)) {
                    foreach (var part in style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                        var pieces = part.Split(new[] { ':' }, 2);
                        if (pieces.Length == 2 && pieces[0].Trim().Equals("width", StringComparison.OrdinalIgnoreCase)) {
                            width = pieces[1].Trim();
                            break;
                        }
                    }
                }
                width ??= col.GetAttribute("width");
                if (string.IsNullOrEmpty(width)) {
                    continue;
                }

                int span = 1;
                if (int.TryParse(col.GetAttribute("span"), out int sp) && sp > 1) {
                    span = sp;
                }

                int size;
                TableWidthUnitValues thisType;
                if (width.EndsWith("%", StringComparison.Ordinal)) {
                    var num = width.TrimEnd('%');
                    if (!int.TryParse(num, out int pct)) {
                        continue;
                    }
                    size = pct * 50;
                    thisType = TableWidthUnitValues.Pct;
                } else {
                    var parser = new CssParser();
                    var decl = parser.ParseDeclaration($"x:{width}");
                    if (TryConvertToTwip(decl.GetProperty("x")?.RawValue, out int w)) {
                        size = w;
                        thisType = TableWidthUnitValues.Dxa;
                    } else {
                        continue;
                    }
                }

                if (widthType == null) {
                    widthType = thisType;
                } else if (widthType != thisType) {
                    return; // mixed width types not supported
                }

                for (int i = 0; i < span; i++) {
                    widths.Add(size);
                }
            }

            if (widthType != null && widths.Count == cols) {
                wordTable.ColumnWidth = widths;
                wordTable.ColumnWidthType = widthType;
            }
        }

        private static void ApplyTableStyles(WordTable wordTable, IHtmlTableElement tableElem) {
            var style = tableElem.GetAttribute("style");
            var borderAttr = tableElem.GetAttribute("border");
            if (string.IsNullOrWhiteSpace(style) && string.IsNullOrWhiteSpace(borderAttr)) {
                return;
            }

            string? background = null;
            int? padTop = null, padRight = null, padBottom = null, padLeft = null;
            BorderValues? tableBorderStyle = null;
            UInt32Value? tableBorderSize = null;
            SixColor tableBorderColor = default;
            bool borderSpecified = false;
            bool collapse = true;

            if (!string.IsNullOrWhiteSpace(style)) {
                foreach (var part in style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                    var pieces = part.Split(new[] { ':' }, 2);
                    if (pieces.Length != 2) {
                        continue;
                    }
                    var name = pieces[0].Trim().ToLowerInvariant();
                    var value = pieces[1].Trim();
                    switch (name) {
                        case "border":
                            if (TryParseBorder(value, out var bStyle, out var bSize, out var bColor)) {
                                tableBorderStyle = bStyle;
                                tableBorderSize = bSize;
                                tableBorderColor = bColor;
                                borderSpecified = true;
                            }
                            break;
                        case "background-color":
                            var color = NormalizeColor(value);
                            if (color != null) {
                                background = color;
                            }
                            break;
                        case "border-collapse":
                            if (value.Equals("separate", StringComparison.OrdinalIgnoreCase)) {
                                collapse = false;
                            }
                            break;
                        case "width":
                            if (value.EndsWith("%", StringComparison.Ordinal)) {
                                var num = value.TrimEnd('%');
                                if (int.TryParse(num, out int pct)) {
                                    wordTable.Width = pct * 50;
                                    wordTable.WidthType = TableWidthUnitValues.Pct;
                                }
                            } else {
                                var parser = new CssParser();
                                var decl = parser.ParseDeclaration($"x:{value}");
                                if (TryConvertToTwip(decl.GetProperty("x")?.RawValue, out int w)) {
                                    wordTable.Width = w;
                                    wordTable.WidthType = TableWidthUnitValues.Dxa;
                                }
                            }
                            break;
                        case "padding": {
                                var parser = new CssParser();
                                var decl = parser.ParseDeclaration($"x:{value}");
                                if (TryConvertToTwip(decl.GetProperty("x")?.RawValue, out int p)) padTop = padRight = padBottom = padLeft = p;
                                break;
                            }
                        case "padding-top": {
                                var parser = new CssParser();
                                var decl = parser.ParseDeclaration($"x:{value}");
                                if (TryConvertToTwip(decl.GetProperty("x")?.RawValue, out int pt)) padTop = pt;
                                break;
                            }
                        case "padding-right": {
                                var parser = new CssParser();
                                var decl = parser.ParseDeclaration($"x:{value}");
                                if (TryConvertToTwip(decl.GetProperty("x")?.RawValue, out int pr)) padRight = pr;
                                break;
                            }
                        case "padding-bottom": {
                                var parser = new CssParser();
                                var decl = parser.ParseDeclaration($"x:{value}");
                                if (TryConvertToTwip(decl.GetProperty("x")?.RawValue, out int pb)) padBottom = pb;
                                break;
                            }
                        case "padding-left": {
                                var parser = new CssParser();
                                var decl = parser.ParseDeclaration($"x:{value}");
                                if (TryConvertToTwip(decl.GetProperty("x")?.RawValue, out int pl)) padLeft = pl;
                                break;
                            }
                    }
                }
            }

            if (!borderSpecified && !string.IsNullOrWhiteSpace(borderAttr)) {
                if (TryParseBorderWidth(borderAttr + "px", out var bSize) && bSize != null) {
                    tableBorderStyle = BorderValues.Single;
                    tableBorderSize = bSize;
                    tableBorderColor = SixColor.Black;
                    borderSpecified = true;
                }
            }

            if (borderSpecified && tableBorderStyle.HasValue && tableBorderSize != null) {
                if (collapse) {
                    wordTable.StyleDetails.SetBordersForAllSides(tableBorderStyle.Value, tableBorderSize, tableBorderColor);
                } else {
                    var hex = tableBorderColor.ToHexColor();
                    foreach (var row in wordTable.Rows) {
                        foreach (var cell in row.Cells) {
                            cell.Borders.LeftStyle = cell.Borders.RightStyle = cell.Borders.TopStyle = cell.Borders.BottomStyle = tableBorderStyle;
                            cell.Borders.LeftSize = cell.Borders.RightSize = cell.Borders.TopSize = cell.Borders.BottomSize = tableBorderSize;
                            cell.Borders.LeftColorHex = cell.Borders.RightColorHex = cell.Borders.TopColorHex = cell.Borders.BottomColorHex = hex;
                        }
                    }
                }
            }

            if (background != null) {
                foreach (var row in wordTable.Rows) {
                    foreach (var cell in row.Cells) {
                        cell.ShadingFillColorHex = background;
                    }
                }
            }

            if (padTop != null) wordTable.StyleDetails.MarginDefaultTopWidth = (short)padTop.Value;
            if (padBottom != null) wordTable.StyleDetails.MarginDefaultBottomWidth = (short)padBottom.Value;
            if (padLeft != null) wordTable.StyleDetails.MarginDefaultLeftWidth = (short)padLeft.Value;
            if (padRight != null) wordTable.StyleDetails.MarginDefaultRightWidth = (short)padRight.Value;
        }

        private static void ApplyRowStyles(WordTableRow row, IHtmlTableRowElement htmlRow) {
            var style = htmlRow.GetAttribute("style");
            if (string.IsNullOrWhiteSpace(style)) {
                return;
            }

            string? background = null;
            BorderValues? borderStyle = null;
            UInt32Value? borderSize = null;
            SixColor borderColor = default;

            foreach (var part in style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                var pieces = part.Split(new[] { ':' }, 2);
                if (pieces.Length != 2) {
                    continue;
                }
                var name = pieces[0].Trim().ToLowerInvariant();
                var value = pieces[1].Trim();
                switch (name) {
                    case "background-color":
                        var color = NormalizeColor(value);
                        if (color != null) background = color;
                        break;
                    case "border":
                        if (TryParseBorder(value, out var bStyle, out var bSize, out var bColor)) {
                            borderStyle = bStyle;
                            borderSize = bSize;
                            borderColor = bColor;
                        }
                        break;
                }
            }

            foreach (var cell in row.Cells) {
                if (background != null) {
                    cell.ShadingFillColorHex = background;
                }
                if (borderStyle != null) {
                    cell.Borders.LeftStyle = cell.Borders.RightStyle = cell.Borders.TopStyle = cell.Borders.BottomStyle = borderStyle;
                    if (borderSize != null) {
                        cell.Borders.LeftSize = cell.Borders.RightSize = cell.Borders.TopSize = cell.Borders.BottomSize = borderSize;
                    }
                    var hex = borderColor.ToHexColor();
                    cell.Borders.LeftColorHex = cell.Borders.RightColorHex = cell.Borders.TopColorHex = cell.Borders.BottomColorHex = hex;
                }
            }
        }

        private static JustificationValues? ApplyCellStyles(WordTableCell cell, IHtmlTableCellElement htmlCell) {
            if (htmlCell == null) {
                return null;
            }
            var style = htmlCell.GetAttribute("style");
            var borderAttr = htmlCell.GetAttribute("border");
            var alignAttr = htmlCell.GetAttribute("align");
            if (string.IsNullOrWhiteSpace(style) && string.IsNullOrWhiteSpace(borderAttr) && string.IsNullOrWhiteSpace(alignAttr)) {
                return null;
            }

            JustificationValues? alignment = null;
            bool borderSet = false;
            if (!string.IsNullOrWhiteSpace(style)) {
                foreach (var part in style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                    var pieces = part.Split(new[] { ':' }, 2);
                    if (pieces.Length != 2) {
                        continue;
                    }
                    var name = pieces[0].Trim().ToLowerInvariant();
                    var value = pieces[1].Trim();
                    switch (name) {
                        case "background-color":
                            var color = NormalizeColor(value);
                            if (color != null) cell.ShadingFillColorHex = color;
                            break;
                        case "width":
                            if (value.EndsWith("%", StringComparison.Ordinal)) {
                                var num = value.TrimEnd('%');
                                if (int.TryParse(num, out int pct)) {
                                    cell.Width = pct * 50;
                                    cell.WidthType = TableWidthUnitValues.Pct;
                                }
                            } else {
                                var parser = new CssParser();
                                var decl = parser.ParseDeclaration($"x:{value}");
                                if (TryConvertToTwip(decl.GetProperty("x")?.RawValue, out int w)) {
                                    cell.Width = w;
                                    cell.WidthType = TableWidthUnitValues.Dxa;
                                }
                            }
                            break;
                        case "border":
                            if (TryParseBorder(value, out var bStyle, out var bSize, out var bColor) && bSize != null) {
                                cell.Borders.LeftStyle = cell.Borders.RightStyle = cell.Borders.TopStyle = cell.Borders.BottomStyle = bStyle;
                                cell.Borders.LeftSize = cell.Borders.RightSize = cell.Borders.TopSize = cell.Borders.BottomSize = bSize;
                                var hex = bColor.ToHexColor();
                                cell.Borders.LeftColorHex = cell.Borders.RightColorHex = cell.Borders.TopColorHex = cell.Borders.BottomColorHex = hex;
                                borderSet = true;
                            }
                            break;
                        case "text-align":
                            var align = value.ToLowerInvariant();
                            alignment = align switch {
                                "center" => JustificationValues.Center,
                                "right" => JustificationValues.Right,
                                "justify" => JustificationValues.Both,
                                "left" => JustificationValues.Left,
                                _ => alignment
                            };
                            break;
                    }
                }
            }

            if (alignment == null && !string.IsNullOrWhiteSpace(alignAttr)) {
                var align = alignAttr.Trim().ToLowerInvariant();
                alignment = align switch {
                    "center" => JustificationValues.Center,
                    "right" => JustificationValues.Right,
                    "justify" => JustificationValues.Both,
                    "left" => JustificationValues.Left,
                    _ => alignment
                };
            }

            if (!borderSet && !string.IsNullOrWhiteSpace(borderAttr)) {
                if (TryParseBorderWidth(borderAttr + "px", out var bSize) && bSize != null) {
                    cell.Borders.LeftStyle = cell.Borders.RightStyle = cell.Borders.TopStyle = cell.Borders.BottomStyle = BorderValues.Single;
                    cell.Borders.LeftSize = cell.Borders.RightSize = cell.Borders.TopSize = cell.Borders.BottomSize = bSize;
                    cell.Borders.LeftColorHex = cell.Borders.RightColorHex = cell.Borders.TopColorHex = cell.Borders.BottomColorHex = "000000";
                }
            }
            return alignment;
        }

        private static bool TryParseBorder(string value, out BorderValues style, out UInt32Value? size, out SixColor color) {
            style = BorderValues.Single;
            size = 4U;
            color = SixColor.Black;
            bool found = false;
            foreach (var part in value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                var token = part.Trim().ToLowerInvariant();
                if (TryParseBorderWidth(token, out var s) && s != null) {
                    size = s;
                    found = true;
                } else if (token == "solid" || token == "dotted" || token == "dashed" || token == "double" || token == "none") {
                    style = token switch {
                        "dotted" => BorderValues.Dotted,
                        "dashed" => BorderValues.Dashed,
                        "double" => BorderValues.Double,
                        "none" => BorderValues.None,
                        _ => BorderValues.Single
                    };
                    found = true;
                } else {
                    var hex = NormalizeColor(token);
                    if (hex != null) {
                        color = SixColor.Parse("#" + hex);
                        found = true;
                    }
                }
            }
            return found;
        }

        private static bool TryParseBorderWidth(string token, out UInt32Value? size) {
            size = null;
            var raw = token.Trim().ToLowerInvariant();
            if (raw.EndsWith("px") && double.TryParse(raw.Substring(0, raw.Length - 2), out double px)) {
                size = (UInt32Value)(uint)Math.Max(1, Math.Round(px * 6));
                return true;
            }
            if (raw.EndsWith("pt") && double.TryParse(raw.Substring(0, raw.Length - 2), out double pt)) {
                size = (UInt32Value)(uint)Math.Max(1, Math.Round(pt * 8));
                return true;
            }
            return false;
        }
    }
}