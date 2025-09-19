using System.Text;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Helper shapes for column pagination
    private abstract class ColItem { public string Kind = string.Empty; }
    private sealed class ColPar : ColItem { public RichParagraphBlock Block = null!; public System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines = null!; public System.Collections.Generic.List<double> Heights = null!; public double Leading; public double Size; public ColPar() { Kind = "P"; } }
    private sealed class ColHead : ColItem { public HeadingBlock Block = null!; public System.Collections.Generic.List<string> Lines = null!; public double Leading; public double Size; public ColHead() { Kind = "H"; } }
    private sealed class ColRule : ColItem { public HorizontalRuleBlock Block = null!; public ColRule() { Kind = "R"; } }
    private sealed class ColImg : ColItem { public ImageBlock Block = null!; public ColImg() { Kind = "I"; } }
    private static string BuildFooter(PdfOptions opts, int page, int pages) {
        string text;
        if (opts.FooterSegments != null && opts.FooterSegments.Count > 0) {
            var sbFooter = new StringBuilder();
            foreach (var seg in opts.FooterSegments) {
                switch (seg.Kind) {
                    case FooterSegmentKind.Text: sbFooter.Append(seg.Text); break;
                    case FooterSegmentKind.PageNumber: sbFooter.Append(page.ToString(CultureInfo.InvariantCulture)); break;
                    case FooterSegmentKind.TotalPages: sbFooter.Append(pages.ToString(CultureInfo.InvariantCulture)); break;
                }
            }
            text = sbFooter.ToString();
        } else {
            text = opts.FooterFormat.Replace("{page}", page.ToString(CultureInfo.InvariantCulture)).Replace("{pages}", pages.ToString(CultureInfo.InvariantCulture));
        }
        double width = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        double em = GlyphWidthEmFor(ChooseNormal(opts.FooterFont));
        double textWidth = text.Length * opts.FooterFontSize * em;
        double x = opts.MarginLeft;
        if (opts.FooterAlign == PdfAlign.Center) x = opts.MarginLeft + Math.Max(0, (width - textWidth) / 2);
        else if (opts.FooterAlign == PdfAlign.Right) x = opts.MarginLeft + Math.Max(0, width - textWidth);
        double y = opts.MarginBottom - opts.FooterOffsetY;
        var sb = new StringBuilder();
        sb.Append("BT\n");
        sb.Append("/F1 ").Append(F(opts.FooterFontSize)).Append(" Tf\n");
        sb.Append("1 0 0 1 ").Append(F(x)).Append(' ').Append(F(y)).Append(" Tm\n");
        sb.Append('(').Append(EscapeText(text)).Append(") Tj\n");
        sb.Append("ET\n");
        return sb.ToString();
    }

    private static LayoutResult LayoutBlocks(IEnumerable<IPdfBlock> blocks, PdfOptions opts) {
        double width = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        double yStart = opts.PageHeight - opts.MarginTop;
        double y = yStart;

        var sb = new StringBuilder();
        var pages = new System.Collections.Generic.List<LayoutResult.Page>();
        var currentPage = new LayoutResult.Page();
        bool usedBold = false;
        bool usedItalic = false;
        bool usedBoldItalic = false;

        void FlushPage() { currentPage.Content = sb.ToString(); pages.Add(currentPage); currentPage = new LayoutResult.Page(); sb.Clear(); }
        void NewPage() { FlushPage(); y = yStart; }

        void WriteLinesInternal(string fontRes, double fontSize, double lineHeight, double x, double widthUsed, double startY, System.Collections.Generic.IReadOnlyList<string> lines, PdfAlign align, PdfColor? color = null, bool applyBaselineTweak = false) {
            sb.Append("BT\n");
            sb.Append('/').Append(fontRes).Append(' ').Append(F(fontSize)).Append(" Tf\n");
            sb.Append(F(lineHeight)).Append(" TL\n");
            double yStart2 = startY;
            if (applyBaselineTweak) {
                var font = fontRes == "F2" ? ChooseBold(ChooseNormal(opts.DefaultFont)) : ChooseNormal(opts.DefaultFont);
                yStart2 -= GetDescender(font, fontSize) * 0.0;
            }
            sb.Append("1 0 0 1 ").Append(F(x)).Append(' ').Append(F(yStart2)).Append(" Tm\n");
            var effectiveColor = color ?? opts.DefaultTextColor;
            if (effectiveColor.HasValue) sb.Append(SetFillColor(effectiveColor.Value));
            for (int i = 0; i < lines.Count; i++) {
                string line = lines[i];
                double dx = 0;
                double em = fontRes == "F2" ? GlyphWidthEmFor(ChooseBold(ChooseNormal(opts.DefaultFont))) : GlyphWidthEmFor(ChooseNormal(opts.DefaultFont));
                double estWidth = line.Length * fontSize * em;
                if (align == PdfAlign.Center) dx = Math.Max(0, (widthUsed - estWidth) / 2);
                else if (align == PdfAlign.Right) dx = Math.Max(0, (widthUsed - estWidth));
                if (dx != 0) sb.Append(F(dx)).Append(" 0 Td\n");
                sb.Append('<').Append(EncodeWinAnsiHex(line)).Append("> Tj\n");
                if (dx != 0) sb.Append(F(-dx)).Append(" 0 Td\n");
                if (i != lines.Count - 1) sb.Append("T*\n");
            }
            sb.Append("ET\n");
        }

        void WriteLines(string fontRes, double fontSize, double lineHeight, double x, double startY, System.Collections.Generic.IReadOnlyList<string> lines, PdfAlign align, PdfColor? color = null, bool applyBaselineTweak = false)
            => WriteLinesInternal(fontRes, fontSize, lineHeight, x, width, startY, lines, align, color, applyBaselineTweak);

        double glyphWidthEm = GlyphWidthEmFor(ChooseNormal(opts.DefaultFont));
        foreach (var block in blocks) {
            if (block is PageBreakBlock) { NewPage(); continue; }
            if (block is HeadingBlock hb) {
                double size = hb.Level switch { 1 => 24, 2 => 18, 3 => 14, _ => 12 };
                double leading = size * 1.25;
                var lines = WrapMonospace(hb.Text, width, size, GlyphWidthEmFor(ChooseBold(ChooseNormal(opts.DefaultFont))));
                double needed = lines.Count * leading + leading * 0.25;
                if (y - needed < opts.MarginBottom) { NewPage(); }
                if (!string.IsNullOrEmpty(hb.LinkUri)) {
                    var baseFont = ChooseBold(ChooseNormal(opts.DefaultFont));
                    double asc = GetAscender(baseFont, size);
                    double desc = GetDescender(baseFont, size);
                    double x1 = opts.MarginLeft;
                    double x2 = opts.MarginLeft + Math.Min(width, lines[0].Length * size * GlyphWidthEmFor(baseFont));
                    double y1 = y - leading - desc;
                    double y2 = y - desc + asc;
                    currentPage.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = hb.LinkUri! });
                }
                WriteLines("F2", size, leading, opts.MarginLeft, y, lines, hb.Align, hb.Color, applyBaselineTweak: false);
                y -= needed;
            }
            // no PlainParagraphBlock in current model
            else if (block is RichParagraphBlock rpb) {
                double size = opts.DefaultFontSize;
                double leading = size * 1.4;
                var (lines, lineHeights) = WrapRichRuns(rpb.Runs, width, size, ChooseNormal(opts.DefaultFont));
                double needed = lineHeights.Sum() + leading * 0.3;
                if (y - needed < opts.MarginBottom) { NewPage(); }
                WriteRichParagraph(sb, rpb, lines, lineHeights, opts, y, size, leading, currentPage.Annotations, null, width);
                usedBold |= rpb.Runs.Any(r => r.Bold);
                usedItalic |= rpb.Runs.Any(r => r.Italic);
                usedBoldItalic |= rpb.Runs.Any(r => r.Bold && r.Italic);
                y -= needed;
            }
            else if (block is BulletListBlock bl) {
                double size = opts.DefaultFontSize;
                double leading = size * 1.4;
                foreach (var text in bl.Items) {
                    var lines = WrapMonospace(text, width, size, glyphWidthEm);
                    double needed = lines.Count * leading + leading * 0.15;
                    if (y - needed < opts.MarginBottom) { NewPage(); }
                    WriteLines("F1", size, leading, opts.MarginLeft, y, lines, bl.Align, bl.Color, applyBaselineTweak: true);
                    y -= needed;
                }
            }
            else if (block is TableBlock tb) {
                var style = tb.Style ?? opts.DefaultTableStyle ?? TableStyles.Light();
                int cols = tb.Rows.Count > 0 ? tb.Rows.Max(r => r.Length) : 0;
                if (cols == 0) continue;
                double padX = style.CellPaddingX;
                double padY = style.CellPaddingY;
                double colGapPx = 0;
                double contentWidth = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
                double tableWidth = contentWidth;
                double[] colPixel = new double[cols];
                for (int c = 0; c < cols; c++) colPixel[c] = (tableWidth - (cols - 1) * colGapPx) / cols;
                double size = opts.DefaultFontSize;
                var normalFont = ChooseNormal(opts.DefaultFont);
                double emMono = GlyphWidthEmFor(normalFont);

                var rowHeights = new double[tb.Rows.Count];
                for (int ri = 0; ri < tb.Rows.Count; ri++) {
                    string[] row = tb.Rows[ri];
                    double maxH = size * 1.6;
                    for (int ci = 0; ci < row.Length; ci++) {
                        string cell = row[ci] ?? string.Empty;
                        double charEm = GlyphWidthEmFor(normalFont);
                        int maxChars = Math.Max(1, (int)System.Math.Floor((colPixel[ci] - 2 * padX) / (size * charEm)));
                        int linesCount = Math.Max(1, (int)System.Math.Ceiling((double)cell.Length / System.Math.Max(1, maxChars)));
                        maxH = System.Math.Max(maxH, linesCount * size * 1.4 + 2 * padY);
                    }
                    rowHeights[ri] = maxH;
                }
                double totalHeight = rowHeights.Sum();
                if (y - totalHeight < opts.MarginBottom) { NewPage(); }
                double xOrigin = opts.MarginLeft;
                if (tb.Align == PdfAlign.Center) xOrigin = opts.MarginLeft + System.Math.Max(0, (contentWidth - tableWidth) / 2);
                else if (tb.Align == PdfAlign.Right) xOrigin = opts.MarginLeft + System.Math.Max(0, contentWidth - tableWidth);

                for (int rowIndex = 0; rowIndex < tb.Rows.Count; rowIndex++) {
                    var row = tb.Rows[rowIndex];
                    double rowHeight = rowHeights[rowIndex];
                    double rowTop = y;
                    double rowBottom = y - rowHeight;
                    if (opts.Debug?.ShowTableRowBoxes == true) DrawRowRect(sb, new PdfColor(1, 0, 1), 0.6, xOrigin, rowBottom, tableWidth, rowHeight);
                    if (style?.HeaderFill is not null && rowIndex == 0) DrawRowFill(sb, style.HeaderFill.Value, xOrigin, rowBottom, tableWidth, rowHeight);
                    else if (style?.RowStripeFill is not null && rowIndex % 2 == 1) DrawRowFill(sb, style.RowStripeFill.Value, xOrigin, rowBottom, tableWidth, rowHeight);
                    if (opts.Debug?.ShowTableBaselines == true) {
                        double x1 = xOrigin;
                        double x2 = xOrigin + tableWidth;
                        double baselineYDbg = rowBottom + padY + GetDescender(normalFont, size);
                        DrawHLine(sb, new PdfColor(0, 0.6, 0), 0.4, x1, x2, baselineYDbg);
                    }
                    double yBase = rowBottom + padY + GetDescender(normalFont, size) + (style?.RowBaselineOffset ?? 0);
                    double xi = xOrigin;
                    double yRect = rowBottom;
                    double rowWidth = tableWidth;
                    double hRect = rowHeight;
                    string fontRes = (rowIndex == 0) ? "F2" : "F1";
                    var textColor = (rowIndex == 0 ? style?.HeaderTextColor : style?.TextColor);
                    for (int c = 0; c < cols && c < row.Length; c++) {
                        string cell = row[c] ?? string.Empty;
                        double innerW = colPixel[c] - 2 * padX;
                        double charEm = GlyphWidthEmFor(ChooseNormal(opts.DefaultFont));
                        double textW = cell.Length * size * charEm;
                        var align = PdfColumnAlign.Left;
                        if (style?.Alignments != null && c < style.Alignments.Count) align = style.Alignments[c];
                        if (style?.RightAlignNumeric == true && LooksNumeric(cell)) align = PdfColumnAlign.Right;
                        double offset = 0;
                        if (align == PdfColumnAlign.Center) offset = System.Math.Max(0, (innerW - textW) / 2);
                        else if (align == PdfColumnAlign.Right) offset = System.Math.Max(0, innerW - textW);
                        double xCell = xi + padX + offset;
                        double yCell = yBase;
                        WriteCell(sb, fontRes, size, xCell, yCell, cell, textColor, opts);
                        if (tb.Links.TryGetValue((rowIndex, c), out var uri)) {
                            var baseFont = fontRes == "F2" ? ChooseBold(ChooseNormal(opts.DefaultFont)) : ChooseNormal(opts.DefaultFont);
                            double asc = GetAscender(baseFont, size);
                            double desc = GetDescender(baseFont, size);
                            double x1 = xCell;
                            double x2 = xCell + textW;
                            double y1 = yCell - desc;
                            double y2 = yCell + asc;
                            currentPage.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = uri });
                        }
                        xi += colPixel[c] + colGapPx;
                    }
                    if (style?.BorderColor is not null && style.BorderWidth > 0) {
                        DrawRowRect(sb, style.BorderColor.Value, style.BorderWidth, xOrigin, yRect, rowWidth, hRect);
                        double xi2 = xOrigin;
                        double yTop = yRect + hRect;
                        double yBottom = yRect;
                        for (int c = 0; c < cols - 1; c++) {
                            xi2 += colPixel[c];
                            if (opts.Debug?.ShowTableColumnGuides == true)
                                DrawVLine(sb, new PdfColor(0, 0, 1), System.Math.Max(0.3, style.BorderWidth), xi2, yTop, yBottom);
                            else
                                DrawVLine(sb, style.BorderColor.Value, style.BorderWidth, xi2, yTop, yBottom);
                            xi2 += colGapPx;
                        }
                    }
                    y -= rowHeight;
                }
            }
            else if (block is RowBlock rb) {
                // Column geometry
                double contentWidth = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
                int ncols = rb.Columns.Count;
                double[] colXs = new double[ncols];
                double[] colWs = new double[ncols];
                double xAcc = opts.MarginLeft;
                for (int i = 0; i < ncols; i++) { double w = System.Math.Max(0, contentWidth * (rb.Columns[i].WidthPercent / 100.0)); colXs[i] = xAcc; colWs[i] = w; xAcc += w; }

                // Prepare per-column paginatable items
                var colStates = new System.Collections.Generic.List<(int idx, int line)>(ncols);
                var colItems = new System.Collections.Generic.List<System.Collections.Generic.List<ColItem>>(ncols);
                for (int i = 0; i < ncols; i++) {
                    colStates.Add((0, 0));
                    var items = new System.Collections.Generic.List<ColItem>();
                    foreach (var cb in rb.Columns[i].Blocks) {
                        if (cb is HeadingBlock hb2) {
                            double size = hb2.Level switch { 1 => 24, 2 => 18, 3 => 14, _ => 12 };
                            double leading = size * 1.25;
                            var lines = WrapMonospace(hb2.Text, colWs[i], size, GlyphWidthEmFor(ChooseBold(ChooseNormal(opts.DefaultFont))));
                            items.Add(new ColHead { Block = hb2, Lines = lines, Leading = leading, Size = size });
                        } else if (cb is RichParagraphBlock rpb2) {
                            double size = opts.DefaultFontSize;
                            double leading = size * 1.4;
                            var wrap = WrapRichRuns(rpb2.Runs, colWs[i], size, ChooseNormal(opts.DefaultFont));
                            items.Add(new ColPar { Block = rpb2, Lines = wrap.Lines, Heights = wrap.LineHeights, Leading = leading, Size = size });
                        } else if (cb is HorizontalRuleBlock hr2) {
                            items.Add(new ColRule { Block = hr2 });
                        } else if (cb is ImageBlock ib2) {
                            items.Add(new ColImg { Block = ib2 });
                        }
                    }
                    colItems.Add(items);
                }

                bool AnyRemaining() {
                    for (int i = 0; i < ncols; i++) if (colStates[i].idx < colItems[i].Count) return true; return false;
                }

                while (AnyRemaining()) {
                    double avail = y - opts.MarginBottom;
                    if (avail <= 0.5) { NewPage(); avail = y - opts.MarginBottom; }

                    double maxConsumed = 0;
                    for (int ci = 0; ci < ncols; ci++) {
                        var items = colItems[ci];
                        var (idx, line) = colStates[ci];
                        double xCol = colXs[ci];
                        double wCol = colWs[ci];
                        double yCol = y;
                        double consumed = 0;
                        double remain = avail;
                        while (idx < items.Count && remain > 0.1) {
                            var it = items[idx];
                            if (it is ColPar par) {
                                var pblock = par.Block;
                                var lines = par.Lines;
                                var heights = par.Heights;
                                double leading = par.Leading;
                                double size = par.Size;
                                // find how many lines fit
                                int start = line;
                                int take = 0; double hsum = 0;
                                for (int li2 = start; li2 < lines.Count; li2++) {
                                    double hAdd = heights[li2];
                                    if (hsum + hAdd + (li2 == lines.Count - 1 ? leading * 0.3 : 0) > remain) break;
                                    hsum += hAdd; take++;
                                }
                                if (take == 0) break; // not enough space for even one line -> stop column here
                                // slice and draw
                                var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
                                var sliceHeights = new System.Collections.Generic.List<double>();
                                for (int k = 0; k < take; k++) { sliceLines.Add(lines[start + k]); sliceHeights.Add(heights[start + k]); }
                                WriteRichParagraph(sb, pblock, sliceLines, sliceHeights, opts, yCol, size, leading, currentPage.Annotations, xCol, wCol);
                                yCol -= hsum; remain -= hsum; consumed += hsum; line += take;
                                // add paragraph spacing only if finished
                                if (line >= lines.Count) { double space = leading * 0.3; if (space <= remain) { yCol -= space; remain -= space; consumed += space; } idx++; line = 0; }
                            } else if (it is ColHead ch) {
                                var hb2 = ch.Block;
                                var lines = ch.Lines;
                                double leading = ch.Leading;
                                double size = ch.Size;
                                double needed = lines.Count * leading + leading * 0.25;
                                if (needed > remain && consumed > 0) break; // defer to next page slice
                                if (needed > remain && consumed == 0) { remain = 0; break; }
                                WriteLinesInternal("F2", size, leading, xCol, wCol, yCol, lines, hb2.Align, hb2.Color, applyBaselineTweak: false);
                                yCol -= needed; remain -= needed; consumed += needed; idx++;
                            } else if (it is ColRule cr) {
                                var hr2 = cr.Block;
                                double needed = hr2.SpacingBefore + hr2.SpacingAfter;
                                if (needed > remain && consumed > 0) break;
                                if (needed > remain && consumed == 0) { remain = 0; break; }
                                yCol -= hr2.SpacingBefore;
                                double x1 = xCol, x2 = xCol + wCol, yLine = yCol - hr2.Thickness * 0.5;
                                DrawHLine(sb, hr2.Color, System.Math.Max(0.2, hr2.Thickness), x1, x2, yLine);
                                yCol -= hr2.SpacingAfter; remain -= needed; consumed += needed; idx++;
                            } else if (it is ColImg ciimg) {
                                var ib2 = ciimg.Block;
                                double needed = ib2.Height;
                                if (needed > remain && consumed > 0) break;
                                if (needed > remain && consumed == 0) { remain = 0; break; }
                                double xImg = xCol;
                                if (ib2.Align == PdfAlign.Center) xImg = xCol + System.Math.Max(0, (wCol - ib2.Width) / 2);
                                else if (ib2.Align == PdfAlign.Right) xImg = xCol + System.Math.Max(0, wCol - ib2.Width);
                                currentPage.Images.Add(new PageImage { Data = ib2.Data, X = xImg, Y = yCol - ib2.Height, W = ib2.Width, H = ib2.Height });
                                yCol -= ib2.Height; remain -= ib2.Height; consumed += ib2.Height; idx++;
                            }
                        }
                        colStates[ci] = (idx, line);
                        if (consumed > maxConsumed) maxConsumed = consumed;
                    }

                    if (maxConsumed <= 0.01) { NewPage(); continue; }
                    y -= maxConsumed;
                }
            }
            else if (block is ImageBlock ib) {
                double x = opts.MarginLeft;
                double contentWidth = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
                if (ib.Align == PdfAlign.Center) x = opts.MarginLeft + System.Math.Max(0, (contentWidth - ib.Width) / 2);
                else if (ib.Align == PdfAlign.Right) x = opts.MarginLeft + System.Math.Max(0, contentWidth - ib.Width);
                if (y - ib.Height < opts.MarginBottom) { NewPage(); }
                currentPage.Images.Add(new PageImage { Data = ib.Data, X = x, Y = y - ib.Height, W = ib.Width, H = ib.Height });
                y -= ib.Height;
            }
            else if (block is PanelParagraphBlock ppb) {
                double size = opts.DefaultFontSize;
                double leading = size * 1.4;
                // Compute available inner width for text (panel width minus horizontal padding)
                double contentWidth = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
                double innerWidth = ppb.Style.MaxWidth.HasValue
                    ? System.Math.Min(contentWidth, ppb.Style.MaxWidth.Value)
                    : contentWidth;
                double textWidthAvail = System.Math.Max(1, innerWidth - 2 * ppb.Style.PaddingX);
                var (lines, lineHeights) = WrapRichRuns(ppb.Runs, textWidthAvail, size, ChooseNormal(opts.DefaultFont));
                double panelWidth = innerWidth;
                double xLeft = opts.MarginLeft;
                if (ppb.Style.Align == PdfAlign.Center) xLeft = opts.MarginLeft + System.Math.Max(0, (contentWidth - innerWidth) / 2);
                else if (ppb.Style.Align == PdfAlign.Right) xLeft = opts.MarginLeft + System.Math.Max(0, contentWidth - innerWidth);

                if (ppb.Style.KeepTogether) {
                    double textHeight = lineHeights.Sum();
                    double panelTop = y;
                    double panelBottom = y - (ppb.Style.PaddingY + textHeight + ppb.Style.PaddingY);
                    if (panelBottom < opts.MarginBottom) { NewPage(); panelTop = y; panelBottom = y - (ppb.Style.PaddingY + textHeight + ppb.Style.PaddingY); }
                    if (ppb.Style.Background.HasValue) DrawRowFill(sb, ppb.Style.Background.Value, xLeft, panelBottom, panelWidth, panelTop - panelBottom);
                    if (ppb.Style.BorderColor.HasValue && ppb.Style.BorderWidth > 0) DrawRowRect(sb, ppb.Style.BorderColor.Value, ppb.Style.BorderWidth, xLeft, panelBottom, panelWidth, panelTop - panelBottom);
                    WriteRichParagraph(sb, new RichParagraphBlock(ppb.Runs, ppb.Align, ppb.DefaultColor), lines, lineHeights, opts, panelTop - ppb.Style.PaddingY, size, leading, currentPage.Annotations, xLeft + ppb.Style.PaddingX, textWidthAvail);
                    y = panelBottom;
                } else {
                    // Split panel across pages by line slices
                    int li = 0; bool firstSeg = true;
                    while (li < lines.Count) {
                        double avail = y - opts.MarginBottom;
                        if (avail < 0.5) { NewPage(); firstSeg = false; continue; }
                        double topPad = firstSeg ? ppb.Style.PaddingY : 0;
                        // ensure at least one line fits
                        double minLine = lineHeights[li];
                        if (avail < topPad + minLine) { NewPage(); firstSeg = false; continue; }
                        double roomForText = avail - topPad - ppb.Style.PaddingY;
                        int take = 0; double hsum = 0;
                        for (int k = li; k < lines.Count; k++) {
                            double h = lineHeights[k];
                            if (hsum + h > roomForText) break;
                            hsum += h; take++;
                        }
                        bool lastSeg = (li + take) >= lines.Count;
                        double panelTop = y;
                        double usedBottomPad = ppb.Style.PaddingY;
                        if (!lastSeg && topPad + hsum + usedBottomPad > avail) usedBottomPad = System.Math.Max(0, avail - (topPad + hsum));
                        double panelBottom = y - (topPad + hsum + usedBottomPad);
                        if (ppb.Style.Background.HasValue) DrawRowFill(sb, ppb.Style.Background.Value, xLeft, panelBottom, panelWidth, panelTop - panelBottom);
                        if (ppb.Style.BorderColor.HasValue && ppb.Style.BorderWidth > 0) DrawRowRect(sb, ppb.Style.BorderColor.Value, ppb.Style.BorderWidth, xLeft, panelBottom, panelWidth, panelTop - panelBottom);
                        // draw slice
                        var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
                        var sliceHeights = new System.Collections.Generic.List<double>();
                        for (int k = 0; k < take; k++) { sliceLines.Add(lines[li + k]); sliceHeights.Add(lineHeights[li + k]); }
                        WriteRichParagraph(sb, new RichParagraphBlock(ppb.Runs, ppb.Align, ppb.DefaultColor), sliceLines, sliceHeights, opts, panelTop - topPad, size, leading, currentPage.Annotations, xLeft + ppb.Style.PaddingX, textWidthAvail);
                        y = panelBottom; li += take; firstSeg = false;
                        if (li < lines.Count) { NewPage(); }
                    }
                }
            }
        }

        FlushPage();
        var result = new LayoutResult { UsedBold = usedBold, UsedItalic = usedItalic, UsedBoldItalic = usedBoldItalic };
        foreach (var p in pages) result.Pages.Add(p);
        return result;
    }
}
