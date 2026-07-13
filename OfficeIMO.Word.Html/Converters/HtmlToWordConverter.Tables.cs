using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using SixColor = OfficeIMO.Drawing.OfficeColor;
using System.Globalization;

namespace OfficeIMO.Word.Html {
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
            ApplyCssToElement(tableElem);

            int cols = DetermineTableColumnCount(tableElem, rows, options);
            ValidateTableLimit(options, rows, cols);
            WordParagraph? captionParagraph = null;
            if (caption != null && options.TableCaptionPosition == TableCaptionPosition.Above) {
                captionParagraph = cell != null ? cell.AddParagraph("", true)
                    : currentParagraph != null ? currentParagraph.AddParagraphAfterSelf()
                    : headerFooter != null ? headerFooter.AddParagraph("")
                    : section.AddParagraph("");
                captionParagraph.SetStyleId("Caption");
                var props = ApplyParagraphStyleFromCss(captionParagraph, caption);
                ApplyClassStyle(caption, captionParagraph, options);
                ApplyBidiIfPresent(caption, captionParagraph);
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
            if (captionParagraph != null) {
                wordTable = captionParagraph.AddTableAfter(rows, cols);
            } else if (cell != null) {
                wordTable = cell.AddTable(rows, cols);
            } else if (currentParagraph != null) {
                wordTable = currentParagraph.AddTableAfter(rows, cols);
            } else if (headerFooter != null) {
                wordTable = headerFooter.AddTable(rows, cols);
            } else {
                wordTable = section.AddTable(rows, cols);
            }
            ApplyTableStyles(wordTable, tableElem);
            ApplyColumnGroup(wordTable, tableElem, cols);
            var occupied = new bool[rows, cols];
            int rIndex = 0;
            bool useRawSpanAttributes = options.MaxTableCells.HasValue;

            void HandleRows(IHtmlCollection<IHtmlTableRowElement> htmlRows) {
                var groupRowCount = htmlRows.Length;
                for (int localRowIndex = 0; localRowIndex < groupRowCount; localRowIndex++) {
                    var htmlRow = htmlRows[localRowIndex];
                    ApplyCssToElement(htmlRow);
                    var wordRow = wordTable.Rows[rIndex];
                    ApplyRowStyles(wordRow, htmlRow);
                    int cIndex = 0;
                    for (int c = 0; c < htmlRow.Cells.Length; c++) {
                        while (cIndex < cols && occupied[rIndex, cIndex]) {     
                            cIndex++;
                        }

                        var htmlCell = htmlRow.Cells[c];
                        ApplyCssToElement(htmlCell);
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
                            rowSpan = GetHtmlRowSpan(htmlTableCell, useRawSpanAttributes);
                            colSpan = GetHtmlColumnSpan(htmlTableCell, useRawSpanAttributes);
                            if (rowSpan == 0) {
                                rowSpan = groupRowCount - localRowIndex;
                            }
                        }

                        rowSpan = Math.Max(1, rowSpan);
                        colSpan = Math.Max(1, colSpan);
                        rowSpan = Math.Min(rowSpan, rows - rIndex);
                        colSpan = Math.Min(colSpan, cols - cIndex);

                        if (rowSpan > 1 || colSpan > 1) {
                            wordTable.MergeCells(rIndex, cIndex, rowSpan, colSpan);
                            for (int rr = rIndex; rr < rIndex + rowSpan && rr < rows; rr++) {
                                for (int cc = cIndex; cc < cIndex + colSpan && cc < cols; cc++) {
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
                var headerStartIndex = rIndex;
                HandleRows(tableElem.Head.Rows);
                for (int headerIndex = headerStartIndex; headerIndex < rIndex; headerIndex++) {
                    wordTable.Rows[headerIndex].RepeatHeaderRowAtTheTopOfEachPage = true;
                }
            }
            foreach (var body in tableElem.Bodies) {
                HandleRows(body.Rows);
            }
            if (tableElem.Foot != null) {
                HandleRows(tableElem.Foot.Rows);
                if (tableElem.Foot.Rows.Length > 0) {
                    wordTable.ConditionalFormattingLastRow = true;
                }
            }

            if (caption != null && options.TableCaptionPosition == TableCaptionPosition.Below) {
                WordParagraph captionParagraphBelow;
                if (cell != null) {
                    captionParagraphBelow = cell.AddParagraph("", true);
                } else if (headerFooter != null) {
                    captionParagraphBelow = headerFooter.AddParagraph("");
                } else {
                    var paragraph = new Paragraph();
                    wordTable._table.InsertAfterSelf(paragraph);
                    captionParagraphBelow = new WordParagraph(doc, paragraph);
                }
                captionParagraphBelow.SetStyleId("Caption");
                var propsBelow = ApplyParagraphStyleFromCss(captionParagraphBelow, caption);
                ApplyClassStyle(caption, captionParagraphBelow, options);
                ApplyBidiIfPresent(caption, captionParagraphBelow);
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

        private int DetermineTableColumnCount(IHtmlTableElement tableElem, int rows, HtmlToWordOptions options) {
            var occupied = new HashSet<long>();
            int cols = 0;
            int rowIndex = 0;

            void HandleRows(IHtmlCollection<IHtmlTableRowElement> htmlRows) {
                int groupRowCount = htmlRows.Length;
                for (int localRowIndex = 0; localRowIndex < groupRowCount; localRowIndex++) {
                    var htmlRow = htmlRows[localRowIndex];
                    int columnIndex = 0;
                    for (int cellIndex = 0; cellIndex < htmlRow.Cells.Length; cellIndex++) {
                        while (occupied.Contains(GetTableGridKey(rowIndex, columnIndex))) {
                            columnIndex++;
                        }

                        var htmlCell = htmlRow.Cells[cellIndex] as IHtmlTableCellElement;
                        int rowSpan = GetHtmlRowSpan(htmlCell, options.MaxTableCells.HasValue);
                        int colSpan = GetHtmlColumnSpan(htmlCell, options.MaxTableCells.HasValue);
                        if (rowSpan == 0) {
                            rowSpan = groupRowCount - localRowIndex;
                        }

                        rowSpan = Math.Max(1, Math.Min(rowSpan, rows - rowIndex));
                        cols = Math.Max(cols, columnIndex + colSpan);
                        ValidateTableLimit(options, rows, cols);

                        for (int rr = rowIndex; rr < rowIndex + rowSpan && rr < rows; rr++) {
                            for (int cc = columnIndex; cc < columnIndex + colSpan; cc++) {
                                occupied.Add(GetTableGridKey(rr, cc));
                            }
                        }

                        columnIndex += colSpan;
                    }

                    rowIndex++;
                }
            }

            if (tableElem.Head != null) {
                HandleRows(tableElem.Head.Rows);
            }

            foreach (var body in tableElem.Bodies) {
                HandleRows(body.Rows);
            }

            if (tableElem.Foot != null) {
                HandleRows(tableElem.Foot.Rows);
            }

            return cols;
        }

        private static long GetTableGridKey(int row, int column) {
            return ((long)row << 32) | (uint)column;
        }

        private static int GetHtmlRowSpan(IHtmlTableCellElement? htmlCell, bool useRawAttribute) {
            if (htmlCell == null) {
                return 1;
            }

            if (useRawAttribute &&
                int.TryParse(htmlCell.GetAttribute("rowspan"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var rowSpan) &&
                rowSpan >= 0) {
                return rowSpan;
            }

            return htmlCell.RowSpan;
        }

        private static int GetHtmlColumnSpan(IHtmlTableCellElement? htmlCell, bool useRawAttribute) {
            if (htmlCell == null) {
                return 1;
            }

            if (useRawAttribute &&
                int.TryParse(htmlCell.GetAttribute("colspan"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var colSpan) &&
                colSpan > 0) {
                return colSpan;
            }

            return Math.Max(1, htmlCell.ColumnSpan);
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
                var styleColText = style ?? string.Empty;
                if (!string.IsNullOrEmpty(style)) {
                    foreach (var part in styleColText.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
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
                var widthText = width ?? string.Empty;
                if (TryParsePercentWidth(widthText, out int pctWidth)) {
                    size = pctWidth;
                    thisType = TableWidthUnitValues.Pct;
                } else {
                    var parser = new CssParser();
                    var decl = parser.ParseDeclaration($"x:{widthText}");       
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

                int remainingColumns = cols - widths.Count;
                if (remainingColumns <= 0) {
                    return;
                }

                int widthsToAdd = Math.Min(span, remainingColumns);
                for (int i = 0; i < widthsToAdd; i++) {
                    widths.Add(size);
                }
            }

            if (widthType != null && widths.Count == cols) {
                wordTable.ColumnWidth = widths;
                wordTable.ColumnWidthType = widthType;
            }
        }

        private static bool TryParsePercentWidth(string value, out int width) {
            width = 0;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }
            var trimmed = value.Trim();
            if (!trimmed.EndsWith("%", StringComparison.Ordinal)) {
                return false;
            }
            var num = trimmed.TrimEnd('%');
            if (!double.TryParse(num, NumberStyles.Float, CultureInfo.InvariantCulture, out var pct)) {
                return false;
            }
            width = (int)Math.Round(pct * 50);
            return true;
        }

        private static void ApplyTableStyles(WordTable wordTable, IHtmlTableElement tableElem) {
            var style = tableElem.GetAttribute("style");
            var borderAttr = tableElem.GetAttribute("border");
            var alignAttr = tableElem.GetAttribute("align");
            var cellSpacingAttr = tableElem.GetAttribute("cellspacing");
            if (string.IsNullOrWhiteSpace(style) && string.IsNullOrWhiteSpace(borderAttr) && string.IsNullOrWhiteSpace(alignAttr) && string.IsNullOrWhiteSpace(cellSpacingAttr)) {
                return;
            }

            string? background = null;
            string? marginLeft = null;
            string? marginRight = null;
            int? padTop = null, padRight = null, padBottom = null, padLeft = null;
            BorderValues? tableBorderStyle = null;
            UInt32Value? tableBorderSize = null;
            SixColor tableBorderColor = default;
            var sideBorders = new Dictionary<TableBorderSide, (BorderValues Style, UInt32Value Size, SixColor Color)>();
            bool borderSpecified = false;
            bool collapse = true;
            int? cellSpacing = null;

            if (!string.IsNullOrWhiteSpace(cellSpacingAttr) && TryParseTableCellSpacing(cellSpacingAttr!, htmlAttribute: true, out var attrSpacing)) {
                cellSpacing = attrSpacing;
            }

            if (!string.IsNullOrWhiteSpace(style)) {
                foreach (var part in (style ?? string.Empty).Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
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
                        case "border-left":
                        case "border-right":
                        case "border-top":
                        case "border-bottom":
                            if (TryGetBorderSide(name, out var side) && TryParseBorder(value, out var sideStyle, out var sideSize, out var sideColor)) {
                                sideBorders[side] = (sideStyle, sideSize, sideColor);
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
                        case "border-spacing":
                            if (TryParseTableCellSpacing(value, htmlAttribute: false, out var spacing)) {
                                cellSpacing = spacing;
                            }
                            break;
                        case "margin":
                            ApplyMarginShorthand(value, ref marginLeft, ref marginRight);
                            break;
                        case "margin-left":
                            marginLeft = value;
                            break;
                        case "margin-right":
                            marginRight = value;
                            break;
                        case "width":
                            if (value.Equals("auto", StringComparison.OrdinalIgnoreCase)) {
                                wordTable.WidthType = TableWidthUnitValues.Auto;
                                wordTable.Width = 0;
                            } else if (TryParsePercentWidth(value, out int pctWidth)) {
                                wordTable.Width = pctWidth;
                                wordTable.WidthType = TableWidthUnitValues.Pct;
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

            var alignment = ResolveTableAlignment(alignAttr, marginLeft, marginRight);
            if (alignment.HasValue) {
                wordTable.Alignment = alignment.Value;
            }

            if (!borderSpecified && !string.IsNullOrWhiteSpace(borderAttr)) {
                if (TryParseBorderWidth(borderAttr + "px", out var bSize)) {
                    tableBorderStyle = BorderValues.Single;
                    tableBorderSize = bSize;
                    tableBorderColor = SixColor.Black;
                    borderSpecified = true;
                }
            }

            if (borderSpecified && tableBorderStyle.HasValue && tableBorderSize != null) {
                if (collapse) {
                    wordTable.StyleDetails?.SetBordersForAllSides(tableBorderStyle.Value, tableBorderSize, tableBorderColor);
                } else {
                    var hex = tableBorderColor.ToRgbHex();
                    foreach (var row in wordTable.Rows) {
                        foreach (var cell in row.Cells) {
                            cell.Borders.LeftStyle = cell.Borders.RightStyle = cell.Borders.TopStyle = cell.Borders.BottomStyle = tableBorderStyle;
                            cell.Borders.LeftSize = cell.Borders.RightSize = cell.Borders.TopSize = cell.Borders.BottomSize = tableBorderSize;
                            cell.Borders.LeftColorHex = cell.Borders.RightColorHex = cell.Borders.TopColorHex = cell.Borders.BottomColorHex = hex;
                        }
                    }
                }
            }

            if (sideBorders.Count > 0) {
                var hasTop = sideBorders.TryGetValue(TableBorderSide.Top, out var top);
                var hasBottom = sideBorders.TryGetValue(TableBorderSide.Bottom, out var bottom);
                var hasLeft = sideBorders.TryGetValue(TableBorderSide.Left, out var left);
                var hasRight = sideBorders.TryGetValue(TableBorderSide.Right, out var right);
                wordTable.StyleDetails?.SetCustomBorders(
                    topStyle: hasTop ? top.Style : null,
                    topSize: hasTop ? top.Size : null,
                    topColor: hasTop ? top.Color : null,
                    bottomStyle: hasBottom ? bottom.Style : null,
                    bottomSize: hasBottom ? bottom.Size : null,
                    bottomColor: hasBottom ? bottom.Color : null,
                    leftStyle: hasLeft ? left.Style : null,
                    leftSize: hasLeft ? left.Size : null,
                    leftColor: hasLeft ? left.Color : null,
                    rightStyle: hasRight ? right.Style : null,
                    rightSize: hasRight ? right.Size : null,
                    rightColor: hasRight ? right.Color : null);
            }

            if (background != null) {
                foreach (var row in wordTable.Rows) {
                    foreach (var cell in row.Cells) {
                        cell.ShadingFillColorHex = background;
                    }
                }
            }

            var styleDetails = wordTable.StyleDetails;
            if (styleDetails != null) {
                if (padTop != null) styleDetails.MarginDefaultTopWidth = (short)padTop.Value;
                if (padBottom != null) styleDetails.MarginDefaultBottomWidth = (short)padBottom.Value;
                if (padLeft != null) styleDetails.MarginDefaultLeftWidth = (short)padLeft.Value;
                if (padRight != null) styleDetails.MarginDefaultRightWidth = (short)padRight.Value;
                if (cellSpacing != null) styleDetails.CellSpacing = (short)cellSpacing.Value;
            }
        }

        private static bool TryParseTableCellSpacing(string value, bool htmlAttribute, out int twips) {
            twips = 0;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            var token = value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
            if (string.IsNullOrWhiteSpace(token)) {
                return false;
            }

            if (htmlAttribute && double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out var pixels)) {
                if (pixels < 0) {
                    return false;
                }

                return TryRoundTableCellSpacingTwips(pixels * 15, out twips);
            }

            var parser = new CssParser();
            var decl = parser.ParseDeclaration($"x:{token}");
            if (TryConvertToTwipAllowNegative(decl.GetProperty("x")?.RawValue, out twips)) {
                return twips >= 0 && twips <= short.MaxValue;
            }

            return TryParseTableCellSpacingLengthLiteral(token, out twips);
        }

        private static bool TryParseTableCellSpacingLengthLiteral(string token, out int twips) {
            twips = 0;
            var value = token.Trim().ToLowerInvariant();
            if (value == "0") {
                return true;
            }

            (string Unit, double TwipsPerUnit)[] units = {
                ("rem", _renderDevice.FontSize * 15),
                ("em", _renderDevice.FontSize * 15),
                ("px", 15),
                ("pt", 20),
                ("cm", 1440 / 2.54),
                ("mm", 1440 / 25.4),
                ("in", 1440),
                ("pc", 240),
                ("q", 1440 / 101.6),
            };

            foreach (var (unit, twipsPerUnit) in units) {
                if (!value.EndsWith(unit, StringComparison.Ordinal)) {
                    continue;
                }

                var numberText = value.Substring(0, value.Length - unit.Length);
                if (!double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || number < 0) {
                    return false;
                }

                return TryRoundTableCellSpacingTwips(number * twipsPerUnit, out twips);
            }

            return false;
        }

        private static bool TryRoundTableCellSpacingTwips(double value, out int twips) {
            twips = 0;
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 0 || value > short.MaxValue) {
                return false;
            }

            twips = (int)Math.Round(value);
            return twips <= short.MaxValue;
        }

        private static void ApplyMarginShorthand(string value, ref string? marginLeft, ref string? marginRight) {
            var parts = value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 1) {
                marginLeft = parts[0];
                marginRight = parts[0];
            } else if (parts.Length == 2) {
                marginLeft = parts[1];
                marginRight = parts[1];
            } else if (parts.Length == 3) {
                marginLeft = parts[1];
                marginRight = parts[1];
            } else if (parts.Length >= 4) {
                marginRight = parts[1];
                marginLeft = parts[3];
            }
        }

        private static TableRowAlignmentValues? ResolveTableAlignment(string? alignAttr, string? marginLeft, string? marginRight) {
            if (!string.IsNullOrWhiteSpace(alignAttr)) {
                var attr = alignAttr!.Trim().ToLowerInvariant();
                if (attr == "center" || attr == "middle") {
                    return TableRowAlignmentValues.Center;
                }
                if (attr == "right") {
                    return TableRowAlignmentValues.Right;
                }
                if (attr == "left") {
                    return TableRowAlignmentValues.Left;
                }
            }

            bool leftAuto = IsCssAuto(marginLeft);
            bool rightAuto = IsCssAuto(marginRight);
            if (leftAuto && rightAuto) {
                return TableRowAlignmentValues.Center;
            }
            if (leftAuto) {
                return TableRowAlignmentValues.Right;
            }
            if (rightAuto) {
                return TableRowAlignmentValues.Left;
            }

            return null;
        }

        private static bool IsCssAuto(string? value) =>
            string.Equals(value?.Trim(), "auto", StringComparison.OrdinalIgnoreCase);

        private static void ApplyRowStyles(WordTableRow row, IHtmlTableRowElement htmlRow) {
            var style = htmlRow.GetAttribute("style");
            if (string.IsNullOrWhiteSpace(style)) {
                return;
            }

            string? background = null;
            BorderValues? borderStyle = null;
            UInt32Value? borderSize = null;
            SixColor borderColor = default;
            var sideBorders = new Dictionary<TableBorderSide, (BorderValues Style, UInt32Value Size, SixColor Color)>();

            foreach (var part in (style ?? string.Empty).Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
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
                    case "border-left":
                    case "border-right":
                    case "border-top":
                    case "border-bottom":
                        if (TryGetBorderSide(name, out var side) && TryParseBorder(value, out var sideStyle, out var sideSize, out var sideColor)) {
                            sideBorders[side] = (sideStyle, sideSize, sideColor);
                        }
                        break;
                }
            }

            foreach (var cell in row.Cells) {
                if (background != null) {
                    cell.ShadingFillColorHex = background;
                }
                if (borderStyle != null && borderSize != null) {
                    cell.Borders.LeftStyle = cell.Borders.RightStyle = cell.Borders.TopStyle = cell.Borders.BottomStyle = borderStyle;
                    cell.Borders.LeftSize = cell.Borders.RightSize = cell.Borders.TopSize = cell.Borders.BottomSize = borderSize;
                    var hex = borderColor.ToRgbHex();
                    cell.Borders.LeftColorHex = cell.Borders.RightColorHex = cell.Borders.TopColorHex = cell.Borders.BottomColorHex = hex;
                }
                foreach (var sideBorder in sideBorders) {
                    ApplyCellBorder(cell, sideBorder.Key, sideBorder.Value.Style, sideBorder.Value.Size, sideBorder.Value.Color);
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
            var verticalAlignAttr = htmlCell.GetAttribute("valign");
            if (string.IsNullOrWhiteSpace(style) && string.IsNullOrWhiteSpace(borderAttr) && string.IsNullOrWhiteSpace(alignAttr) && string.IsNullOrWhiteSpace(verticalAlignAttr)) {
                return null;
            }

            JustificationValues? alignment = null;
            bool borderSet = false;
            if (TryMapTableCellVerticalAlignment(verticalAlignAttr, out var attrVerticalAlignment)) {
                cell.VerticalAlignment = attrVerticalAlignment;
            }

            if (!string.IsNullOrWhiteSpace(style)) {
                var styleText = style!;
                foreach (var part in styleText.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
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
                            if (TryParsePercentWidth(value, out int pctWidth)) {
                                cell.Width = pctWidth;
                                cell.WidthType = TableWidthUnitValues.Pct;
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
                            if (TryParseBorder(value, out var bStyle, out var bSize, out var bColor)) {
                                cell.Borders.LeftStyle = cell.Borders.RightStyle = cell.Borders.TopStyle = cell.Borders.BottomStyle = bStyle;
                                cell.Borders.LeftSize = cell.Borders.RightSize = cell.Borders.TopSize = cell.Borders.BottomSize = bSize;
                                var hex = bColor.ToRgbHex();
                                cell.Borders.LeftColorHex = cell.Borders.RightColorHex = cell.Borders.TopColorHex = cell.Borders.BottomColorHex = hex;
                                borderSet = true;
                            }
                            break;
                        case "border-left":
                        case "border-right":
                        case "border-top":
                        case "border-bottom":
                            if (TryGetBorderSide(name, out var side) && TryParseBorder(value, out var sideStyle, out var sideSize, out var sideColor)) {
                                ApplyCellBorder(cell, side, sideStyle, sideSize, sideColor);
                                borderSet = true;
                            }
                            break;
                        case "text-align":
                            if (TryMapTextAlign(value, GetBidiFromDir(htmlCell), out var mappedAlignment)) {
                                alignment = mappedAlignment;
                            }
                            break;
                        case "vertical-align":
                            if (TryMapTableCellVerticalAlignment(value, out var verticalAlignment)) {
                                cell.VerticalAlignment = verticalAlignment;
                            }
                            break;
                    }
                }
            }

            if (alignment == null && !string.IsNullOrWhiteSpace(alignAttr)) {
                var align = (alignAttr ?? string.Empty).Trim().ToLowerInvariant();
                alignment = align switch {
                    "center" => JustificationValues.Center,
                    "right" => JustificationValues.Right,
                    "justify" => JustificationValues.Both,
                    "left" => JustificationValues.Left,
                    _ => alignment
                };
            }

            if (!borderSet && !string.IsNullOrWhiteSpace(borderAttr)) {
                if (TryParseBorderWidth(borderAttr + "px", out var bSize)) {
                    cell.Borders.LeftStyle = cell.Borders.RightStyle = cell.Borders.TopStyle = cell.Borders.BottomStyle = BorderValues.Single;
                    cell.Borders.LeftSize = cell.Borders.RightSize = cell.Borders.TopSize = cell.Borders.BottomSize = bSize;
                    cell.Borders.LeftColorHex = cell.Borders.RightColorHex = cell.Borders.TopColorHex = cell.Borders.BottomColorHex = "000000";
                }
            }
            return alignment;
        }

        private static bool TryMapTableCellVerticalAlignment(string? value, out TableVerticalAlignmentValues alignment) {
            alignment = TableVerticalAlignmentValues.Top;
            var normalized = value?.Trim().ToLowerInvariant();
            switch (normalized) {
                case "top":
                    alignment = TableVerticalAlignmentValues.Top;
                    return true;
                case "middle":
                case "center":
                    alignment = TableVerticalAlignmentValues.Center;
                    return true;
                case "bottom":
                    alignment = TableVerticalAlignmentValues.Bottom;
                    return true;
                default:
                    return false;
            }
        }

        private enum TableBorderSide {
            Left,
            Right,
            Top,
            Bottom
        }

        private static bool TryGetBorderSide(string propertyName, out TableBorderSide side) {
            switch (propertyName.ToLowerInvariant()) {
                case "border-left":
                    side = TableBorderSide.Left;
                    return true;
                case "border-right":
                    side = TableBorderSide.Right;
                    return true;
                case "border-top":
                    side = TableBorderSide.Top;
                    return true;
                case "border-bottom":
                    side = TableBorderSide.Bottom;
                    return true;
                default:
                    side = TableBorderSide.Left;
                    return false;
            }
        }

        private static void ApplyCellBorder(WordTableCell cell, TableBorderSide side, BorderValues style, UInt32Value size, SixColor color) {
            var hex = color.ToRgbHex();
            switch (side) {
                case TableBorderSide.Left:
                    cell.Borders.LeftStyle = style;
                    cell.Borders.LeftSize = size;
                    cell.Borders.LeftColorHex = hex;
                    break;
                case TableBorderSide.Right:
                    cell.Borders.RightStyle = style;
                    cell.Borders.RightSize = size;
                    cell.Borders.RightColorHex = hex;
                    break;
                case TableBorderSide.Top:
                    cell.Borders.TopStyle = style;
                    cell.Borders.TopSize = size;
                    cell.Borders.TopColorHex = hex;
                    break;
                case TableBorderSide.Bottom:
                    cell.Borders.BottomStyle = style;
                    cell.Borders.BottomSize = size;
                    cell.Borders.BottomColorHex = hex;
                    break;
            }
        }

        private static bool TryParseBorder(string value, out BorderValues style, out UInt32Value size, out SixColor color) {
            style = BorderValues.Single;
            size = 4U;
            color = SixColor.Black;
            bool found = false;
            foreach (var part in value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                var token = part.Trim().ToLowerInvariant();
                if (TryParseBorderWidth(token, out var s)) {
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

        private static bool TryParseBorderWidth(string token, out UInt32Value size) {
            size = 0;
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
