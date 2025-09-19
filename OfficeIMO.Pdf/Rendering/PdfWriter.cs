using System.Text;
using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfWriter {
    private static readonly char[] WordSplitChars = new[] { ' ', '\n', '\t' };

    public static byte[] Write(PdfDoc doc, IEnumerable<IPdfBlock> blocks, PdfOptions opts, string? title, string? author, string? subject, string? keywords) {
        // Layout blocks into pages and create per-page content streams.
        var layout = LayoutBlocks(blocks, opts);

        // Build PDF objects as byte arrays, then assemble with xref.
        var objects = new List<byte[]>();

        // Reserve IDs (1-based). We'll assign as we add to `objects`.
        int infoId = 0, catalogId = 0, pagesId = 0;
        var pageIds = new List<int>();
        var contentIds = new List<int>();

        // Collect required fonts across all pages (basic: Courier + Courier-Bold if headings used)
        bool needsBold = layout.UsedBold;
        var fonts = new List<FontRef>();
        int fontNormalId = 0;
        int fontBoldId = 0;

        // Add font objects
        var baseFont = ChooseNormal(opts.DefaultFont);
        fontNormalId = AddObject(objects, "<< /Type /Font /Subtype /Type1 /BaseFont /" + baseFont.ToBaseFontName() + " /Encoding /WinAnsiEncoding >>\n");
        fonts.Add(new FontRef("F1", baseFont, fontNormalId));
        if (needsBold) {
            var boldFont = ChooseBold(baseFont);
            fontBoldId = AddObject(objects, "<< /Type /Font /Subtype /Type1 /BaseFont /" + boldFont.ToBaseFontName() + " /Encoding /WinAnsiEncoding >>\n");
            fonts.Add(new FontRef("F2", boldFont, fontBoldId));
        }

        // Create content streams and page objects
        int totalPages = layout.Pages.Count;
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++) {
            var page = layout.Pages[pageIndex];
            // Make a resources dict that references the fonts we declared
            string fontDict = needsBold
                ? $"<< /F1 {fontNormalId} 0 R /F2 {fontBoldId} 0 R >>"
                : $"<< /F1 {fontNormalId} 0 R >>";

            // Content stream
            string contentStr = page.Content;
            if (opts.ShowPageNumbers) {
                string footer = BuildFooter(opts, pageIndex + 1, totalPages);
                contentStr += footer;
            }
            byte[] content = Encoding.ASCII.GetBytes(contentStr);
            int contentId = AddObject(objects, "<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
            // Append raw content bytes + endstream/endobj
            // We'll append extra to the last object content after we compute indices; here we simply merge bytes.
            // For simplicity, rebuild the last object with full content now.
            objects[contentId - 1] = Merge(
                Encoding.ASCII.GetBytes(contentId.ToString(CultureInfo.InvariantCulture) + " 0 obj\n<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n"),
                content,
                Encoding.ASCII.GetBytes("\nendstream\nendobj\n")
            );
            contentIds.Add(contentId);

            // Page object
            int pageId = AddObject(objects,
                "<< /Type /Page /Parent 0 0 R /MediaBox [0 0 " + F0(opts.PageWidth) + " " + F0(opts.PageHeight) + 
                "] /Resources << /Font " + fontDict + " >> /Contents " + contentId.ToString(CultureInfo.InvariantCulture) + " 0 R >>\n");
            pageIds.Add(pageId);
        }

        // Pages tree
        string kids = string.Join(" ", pageIds.Select(id => id.ToString(CultureInfo.InvariantCulture) + " 0 R"));
        pagesId = AddObject(objects, "<< /Type /Pages /Count " + pageIds.Count.ToString(CultureInfo.InvariantCulture) + " /Kids [ " + kids + " ] >>\n");

        // Fix Parent references in each page now that we know pagesId.
        for (int i = 0; i < pageIds.Count; i++) {
            int pageId = pageIds[i];
            string original = Encoding.ASCII.GetString(objects[pageId - 1]);
            string fixedObj = original.Replace("/Parent 0 0 R", "/Parent " + pagesId.ToString(CultureInfo.InvariantCulture) + " 0 R");
            objects[pageId - 1] = Encoding.ASCII.GetBytes(fixedObj);
        }

        // Catalog
        catalogId = AddObject(objects, "<< /Type /Catalog /Pages " + pagesId.ToString(CultureInfo.InvariantCulture) + " 0 R >>\n");

        // Info (metadata)
        var info = new StringBuilder("<< ");
        if (!string.IsNullOrEmpty(title)) info.Append("/Title ").Append(PdfString(title!)).Append(' ');
        if (!string.IsNullOrEmpty(author)) info.Append("/Author ").Append(PdfString(author!)).Append(' ');
        if (!string.IsNullOrEmpty(subject)) info.Append("/Subject ").Append(PdfString(subject!)).Append(' ');
        if (!string.IsNullOrEmpty(keywords)) info.Append("/Keywords ").Append(PdfString(keywords!)).Append(' ');
        info.Append("/Producer (OfficeIMO.Pdf) >>\n");
        infoId = AddObject(objects, info.ToString());

        // Assemble final PDF
        using var ms = new MemoryStream();
        var header = Encoding.ASCII.GetBytes("%PDF-1.4\n%\u00e2\u00e3\u00cf\u00d3\n"); // binary line ensures binary file
        ms.Write(header, 0, header.Length);

        // Write objects and record offsets
        var offsets = new List<long> { 0L }; // index 0 is free object
        for (int i = 0; i < objects.Count; i++) {
            long off = ms.Position;
            offsets.Add(off);
            ms.Write(objects[i], 0, objects[i].Length);
        }

        long xrefPos = ms.Position;
        var sw = new StreamWriter(ms, Encoding.ASCII, 1024, leaveOpen: true) { NewLine = "\n" };
        sw.WriteLine("xref");
        sw.WriteLine("0 " + (objects.Count + 1).ToString(CultureInfo.InvariantCulture));
        sw.WriteLine("0000000000 65535 f ");
        for (int i = 1; i <= objects.Count; i++) {
            sw.WriteLine(offsets[i].ToString("0000000000", CultureInfo.InvariantCulture) + " 00000 n ");
        }
        sw.WriteLine("trailer");
        sw.WriteLine("<< /Size " + (objects.Count + 1).ToString(CultureInfo.InvariantCulture) + " /Root " + catalogId.ToString(CultureInfo.InvariantCulture) + " 0 R /Info " + infoId.ToString(CultureInfo.InvariantCulture) + " 0 R >>");
        sw.WriteLine("startxref");
        sw.WriteLine(xrefPos.ToString(System.Globalization.CultureInfo.InvariantCulture));
        sw.WriteLine("%%EOF");
        sw.Flush();

        return ms.ToArray();
    }

    private static int AddObject(List<byte[]> list, string body) {
        int id = list.Count + 1;
        var bytes = Encoding.ASCII.GetBytes(id.ToString(CultureInfo.InvariantCulture) + " 0 obj\n" + body + "endobj\n");
        list.Add(bytes);
        return id;
    }

    private static byte[] Merge(params byte[][] arrays) {
        int len = arrays.Sum(a => a.Length);
        var buf = new byte[len];
        int pos = 0;
        foreach (var a in arrays) { Buffer.BlockCopy(a, 0, buf, pos, a.Length); pos += a.Length; }
        return buf;
    }

    private static string PdfString(string s) {
        // Literal string in parentheses with robust escaping (incl. control chars via octal)
        return "(" + EscapeLiteral(s) + ")";
    }

    private sealed class LayoutResult {
        public List<Page> Pages { get; } = new();
        public bool UsedBold { get; set; }
        public sealed class Page { public string Content { get; set; } = string.Empty; }
    }

    private static string BuildFooter(PdfOptions opts, int page, int pages) {
        string text = opts.FooterFormat.Replace("{page}", page.ToString(CultureInfo.InvariantCulture)).Replace("{pages}", pages.ToString(CultureInfo.InvariantCulture));
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
        var pages = new List<string>();
        bool usedBold = false;

        void FlushPage() { pages.Add(sb.ToString()); sb.Clear(); }
        // first page is implicit; content buffer starts empty
        void NewPage() { FlushPage(); y = yStart; }

        // Helper to write a text block at (x, y) with leading and multiple lines
        void WriteLines(string fontRes, double fontSize, double lineHeight, double x, double startY, IReadOnlyList<string> lines, PdfAlign align, PdfColor? color = null, bool applyBaselineTweak = false) {
            // Begin text object once per block
            sb.Append("BT\n");
            sb.Append('/').Append(fontRes).Append(' ').Append(F(fontSize)).Append(" Tf\n");
            sb
                .Append(F(lineHeight)).Append(" TL\n");
            double yStart = startY;
            if (applyBaselineTweak) {
                var font = fontRes == "F2" ? ChooseBold(ChooseNormal(opts.DefaultFont)) : ChooseNormal(opts.DefaultFont);
                // Keep baseline close to visual center without altering line height semantics.
                // Using a very small descender-based tweak; can be tuned later if needed.
                yStart -= GetDescender(font, fontSize) * 0.0;
            }
            sb.Append("1 0 0 1 ").Append(F(opts.MarginLeft)).Append(' ').Append(F(yStart)).Append(" Tm\n");
            var effectiveColor = color ?? opts.DefaultTextColor;
            if (effectiveColor.HasValue) sb.Append(SetFillColor(effectiveColor.Value));
            // First line
            for (int i = 0; i < lines.Count; i++) {
                if (i != 0) sb.Append("T*\n");
                var line = lines[i];
                double em = fontRes == "F2" ? 0.6 : 0.6; // same approx for now
                double widthContent = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
                double lineWidth = line.Length * fontSize * em;
                double dx = 0;
                if (align == PdfAlign.Center) dx = Math.Max(0, (widthContent - lineWidth) / 2);
                else if (align == PdfAlign.Right) dx = Math.Max(0, (widthContent - lineWidth));
                if (dx != 0) sb.Append(F(dx)).Append(" 0 Td\n");
                // Emit as hex string encoded in WinAnsi so extended chars (e.g., bullet) render correctly
                sb.Append('<').Append(EncodeWinAnsiHex(line)).Append("> Tj\n");
                if (dx != 0) sb.Append(F(-dx)).Append(" 0 Td\n");
            }
            sb.Append("ET\n");
        }

        // Choose Courier metrics for wrapping predictability
        double glyphWidthEm = GlyphWidthEmFor(ChooseNormal(opts.DefaultFont));

        foreach (var block in blocks) {
            if (block is PageBreakBlock) { NewPage(); continue; }

            if (block is HeadingBlock hb) {
                usedBold = true;
                double size = hb.Level switch { 1 => 24, 2 => 18, 3 => 14, _ => 12 };
                double leading = size * 1.25;
                var lines = WrapMonospace(hb.Text, width, size, GlyphWidthEmFor(ChooseBold(ChooseNormal(opts.DefaultFont))));
                // Page breaks as needed
                double needed = lines.Count * leading + leading * 0.25; // small extra breathing room
                if (y - needed < opts.MarginBottom) { NewPage(); }
                WriteLines("F2", size, leading, opts.MarginLeft, y, lines, hb.Align, hb.Color, applyBaselineTweak: false);
                y -= needed;
            }
            else if (block is ParagraphBlock pb) {
                double size = opts.DefaultFontSize;
                double leading = size * 1.4;
                var lines = WrapMonospace(pb.Text, width, size, glyphWidthEm);
                double needed = lines.Count * leading + leading * 0.3;
                if (y - needed < opts.MarginBottom) { NewPage(); }
                WriteLines("F1", size, leading, opts.MarginLeft, y, lines, pb.Align, pb.Color, applyBaselineTweak: false);
                y -= needed;
            }
            else if (block is BulletListBlock bl) {
                double size = opts.DefaultFontSize;
                double leading = size * 1.4;
                foreach (var item in bl.Items) {
                    var text = "• " + item;
                    var lines = WrapMonospace(text, width, size, glyphWidthEm);
                    double needed = lines.Count * leading;
                    if (y - needed < opts.MarginBottom) { NewPage(); }
                    // Apply the same internal baseline tweak used to align table rows
                    WriteLines("F1", size, leading, opts.MarginLeft, y, lines, bl.Align, bl.Color, applyBaselineTweak: true);
                    y -= needed;
                }
            }
            else if (block is TableBlock tb) {
                // Header row renders with F2; mark bold as used so font gets registered in resources
                if (tb.Rows.Count > 0) usedBold = true;
                double size = opts.DefaultFontSize;
                double leading = size * 1.3;
                // Compute column widths (characters) in monospaced metrics
                int cols = tb.Rows.Count > 0 ? tb.Rows[0].Length : 0;
                var colWidths = new int[cols];
                foreach (var r in tb.Rows) {
                    for (int c = 0; c < cols && c < r.Length; c++) colWidths[c] = Math.Max(colWidths[c], r[c]?.Length ?? 0);
                }
                var style = tb.Style ?? opts.DefaultTableStyle;
                // Geometry
                var normalFont = ChooseNormal(opts.DefaultFont);
                double emMono = GlyphWidthEmFor(normalFont);
                double colGapPx = 0; // rely on CellPadding for spacing; avoids misalignment
                var colPixel = new double[cols];
                for (int c = 0; c < cols; c++) colPixel[c] = colWidths[c] * size * emMono + (style?.CellPaddingX ?? 0) * 2;
                double rowWidth = 0; for (int c = 0; c < cols; c++) rowWidth += colPixel[c]; rowWidth += Math.Max(0, cols - 1) * colGapPx;

                for (int rowIndex = 0; rowIndex < tb.Rows.Count; rowIndex++) {
                    var row = tb.Rows[rowIndex];
                    // Page break check BEFORE drawing the row
                    double needed = leading;
                    if (y - needed < opts.MarginBottom) { NewPage(); }

                    // Optional fills (header / zebra)
                    double xOrigin = opts.MarginLeft;
                    if (tb.Align == PdfAlign.Center) xOrigin += Math.Max(0, (width - rowWidth) / 2);
                    else if (tb.Align == PdfAlign.Right) xOrigin += Math.Max(0, width - rowWidth);
                    double padX = style?.CellPaddingX ?? 0;
                    double padY = style?.CellPaddingY ?? 0;
                    // Row geometry: bottom = current top minus full leading (row height)
                    double rowBottom = y - leading;
                    double yRect = rowBottom;        // draw fills/borders over the full row height
                    double hRect = leading;
                    if (style is not null) {
                        if (rowIndex == 0 && style.HeaderFill.HasValue) DrawRowFill(sb, style.HeaderFill.Value, xOrigin, yRect, rowWidth, hRect);
                        else if (rowIndex % 2 == 1 && style.RowStripeFill.HasValue) DrawRowFill(sb, style.RowStripeFill.Value, xOrigin, yRect, rowWidth, hRect);
                    }
                    var textColor = style?.TextColor;
                    if (rowIndex == 0 && style?.HeaderTextColor is not null) textColor = style.HeaderTextColor;
                    // Header uses bold font
                    string fontRes = rowIndex == 0 ? "F2" : "F1";
                    // Render each cell precisely at computed x positions
                    double xi = xOrigin;
                    // Draw debug overlays if requested
                    if (opts.Debug?.ShowTableRowBoxes == true) DrawRowRect(sb, new PdfColor(1, 0, 1), 0.6, xOrigin, yRect, rowWidth, hRect);
                    if (opts.Debug?.ShowTableBaselines == true) {
                        double x1 = xOrigin; double x2 = xOrigin + rowWidth;
                        double baselineYDbg = rowBottom + padY + GetDescender(normalFont, size) + (style?.RowBaselineOffset ?? 0);
                        DrawHLine(sb, new PdfColor(0, 0.6, 0), 0.4, x1, x2, baselineYDbg);
                    }
                    // Baseline: bottom + padY + descender + optional style offset
                    double yBase = rowBottom + padY + GetDescender(normalFont, size) + (style?.RowBaselineOffset ?? 0);
                    for (int c = 0; c < cols && c < row.Length; c++) {
                        string cell = row[c] ?? string.Empty;
                        // alignment within cell
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
                        double yCell = yBase; // baseline
                        WriteCell(sb, fontRes, size, xCell, yCell, cell, textColor, opts);
                        xi += colPixel[c] + colGapPx;
                    }
                    // Borders (row rect + vertical grid)
                    if (style?.BorderColor is not null && style.BorderWidth > 0) {
                        DrawRowRect(sb, style.BorderColor.Value, style.BorderWidth, xOrigin, yRect, rowWidth, hRect);
                        // Vertical grid lines
                        double xi2 = xOrigin;
                    double yTop = yRect + hRect;
                    double yBottom = yRect;
                    for (int c = 0; c < cols - 1; c++) {
                        xi2 += colPixel[c];
                        if (opts.Debug?.ShowTableColumnGuides == true)
                            DrawVLine(sb, new PdfColor(0, 0, 1), Math.Max(0.3, style.BorderWidth), xi2, yTop, yBottom);
                        else
                            DrawVLine(sb, style.BorderColor.Value, style.BorderWidth, xi2, yTop, yBottom);
                        xi2 += colGapPx;
                    }
                    }
                    y -= needed;
                }
            }
            else {
                // Unknown future block types – skip for now
            }

            // If we've run out of space exactly, force a new page for the next block
            if (y <= opts.MarginBottom + 4) NewPage();
        }

        // Close last page
        FlushPage();

        var result = new LayoutResult { UsedBold = usedBold };
        foreach (var c in pages) result.Pages.Add(new LayoutResult.Page { Content = c });
        return result;
    }

    private static string EscapeText(string s) => EscapeLiteral(s);

    private static string EscapeLiteral(string s) {
        if (string.IsNullOrEmpty(s)) return string.Empty;
        var sb = new StringBuilder(s.Length + 8);
        for (int i = 0; i < s.Length; i++) {
            char ch = s[i];
            switch (ch) {
                case '\\': sb.Append("\\\\"); break;
                case '(': sb.Append("\\("); break;
                case ')': sb.Append("\\)"); break;
                case '\r': sb.Append("\\r"); break;
                case '\n': sb.Append("\\n"); break;
                case '\t': sb.Append("\\t"); break;
                case '\b': sb.Append("\\b"); break;
                case '\f': sb.Append("\\f"); break;
                default:
                    // Escape other control chars as octal to avoid injecting PDF control sequences
                    if (ch < 32 || ch == 127) {
                        int v = ch;
                        // 3-digit octal per PDF spec
                        sb.Append('\\')
                          .Append(((v >> 6) & 0x7).ToString(System.Globalization.CultureInfo.InvariantCulture))
                          .Append(((v >> 3) & 0x7).ToString(System.Globalization.CultureInfo.InvariantCulture))
                          .Append((v & 0x7).ToString(System.Globalization.CultureInfo.InvariantCulture));
                    } else {
                        sb.Append(ch);
                    }
                    break;
            }
        }
        return sb.ToString();
    }

    private static string EncodeWinAnsiHex(string s) {
        var bytes = PdfWinAnsiEncoding.Encode(s);
        var sb = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) sb.Append(bytes[i].ToString("X2", CultureInfo.InvariantCulture));
        return sb.ToString();
    }

    private static string SetFillColor(PdfColor color) => F(color.R) + " " + F(color.G) + " " + F(color.B) + " rg\n";
    private static string SetStrokeColor(PdfColor color) => F(color.R) + " " + F(color.G) + " " + F(color.B) + " RG\n";

    private static void DrawRowFill(StringBuilder sb, PdfColor color, double x, double y, double w, double h) {
        sb.Append("q\n");
        sb.Append(SetFillColor(color));
        sb.Append(F(x)).Append(' ').Append(F(y)).Append(' ').Append(F(w)).Append(' ').Append(F(h)).Append(" re f\n");
        sb.Append("Q\n");
    }

    private static void DrawRowRect(StringBuilder sb, PdfColor color, double widthStroke, double x, double y, double w, double h) {
        sb.Append("q\n");
        sb.Append(SetStrokeColor(color));
        sb.Append(F(widthStroke)).Append(" w\n");
        sb.Append(F(x)).Append(' ').Append(F(y)).Append(' ').Append(F(w)).Append(' ').Append(F(h)).Append(" re S\n");
        sb.Append("Q\n");
    }

    private static void DrawVLine(StringBuilder sb, PdfColor color, double widthStroke, double x, double yTop, double yBottom) {
        sb.Append("q\n");
        sb.Append(SetStrokeColor(color));
        sb.Append(F(widthStroke)).Append(" w\n");
        sb.Append(F(x)).Append(' ').Append(F(yTop)).Append(" m ").Append(F(x)).Append(' ').Append(F(yBottom)).Append(" l S\n");
        sb.Append("Q\n");
    }

    private static void DrawHLine(StringBuilder sb, PdfColor color, double widthStroke, double x1, double x2, double y) {
        sb.Append("q\n");
        sb.Append(SetStrokeColor(color));
        sb.Append(F(widthStroke)).Append(" w\n");
        sb.Append(F(x1)).Append(' ').Append(F(y)).Append(" m ").Append(F(x2)).Append(' ').Append(F(y)).Append(" l S\n");
        sb.Append("Q\n");
    }

    private static void WriteCell(StringBuilder sb, string fontRes, double fontSize, double x, double y, string text, PdfColor? color, PdfOptions opts) {
        sb.Append("BT\n");
        sb.Append('/').Append(fontRes).Append(' ').Append(F(fontSize)).Append(" Tf\n");
        var effective = color ?? opts.DefaultTextColor;
        if (effective.HasValue) sb.Append(SetFillColor(effective.Value));
        sb.Append("1 0 0 1 ").Append(F(x)).Append(' ').Append(F(y)).Append(" Tm\n");
        sb.Append('<').Append(EncodeWinAnsiHex(text)).Append("> Tj\n");
        sb.Append("ET\n");
    }

    private static List<string> WrapMonospace(string text, double widthPts, double fontSize, double glyphWidthEm) {
        // Estimated chars per line for monospaced font
        double glyphWidth = fontSize * glyphWidthEm;
        int maxChars = Math.Max(8, (int)Math.Floor(widthPts / glyphWidth));
        var words = text.Replace("\r", "").Split(WordSplitChars, StringSplitOptions.None);
        var lines = new List<string>();
        var line = new StringBuilder();
        foreach (var w in words) {
            if (w.Contains('\n')) {
                // Should not happen due to split, but keep safe
            }
            if (line.Length == 0) {
                if (w.Length <= maxChars) line.Append(w);
                else {
                    // very long word: hard break
                    for (int i = 0; i < w.Length; i += maxChars) {
                        var chunk = w.Substring(i, Math.Min(maxChars, w.Length - i));
                        lines.Add(chunk);
                    }
                }
            } else {
                if (line.Length + 1 + w.Length <= maxChars) {
                    line.Append(' ').Append(w);
                } else {
                    lines.Add(line.ToString());
                    line.Clear();
                    line.Append(w);
                }
            }
        }
        if (line.Length > 0) lines.Add(line.ToString());
        if (lines.Count == 0) lines.Add(string.Empty);
        return lines;
    }

    private static string F(double d) => d.ToString("0.###", CultureInfo.InvariantCulture);
    private static string F0(double d) => d.ToString("0", CultureInfo.InvariantCulture);

    private static PdfStandardFont ChooseNormal(PdfStandardFont requested) => requested switch {
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold => PdfStandardFont.Helvetica,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold => PdfStandardFont.TimesRoman,
        PdfStandardFont.Courier or PdfStandardFont.CourierBold => PdfStandardFont.Courier,
        _ => PdfStandardFont.Courier
    };

    private static PdfStandardFont ChooseBold(PdfStandardFont normal) => normal switch {
        PdfStandardFont.Helvetica => PdfStandardFont.HelveticaBold,
        PdfStandardFont.TimesRoman => PdfStandardFont.TimesBold,
        PdfStandardFont.Courier => PdfStandardFont.CourierBold,
        _ => PdfStandardFont.CourierBold
    };

    private static double GlyphWidthEmFor(PdfStandardFont font) => font switch {
        PdfStandardFont.Courier or PdfStandardFont.CourierBold => 0.6,
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold => 0.55,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold => 0.5,
        _ => 0.6
    };

    // Approximate descender height (distance from baseline down to glyph bottom) in points.
    private static double GetDescender(PdfStandardFont font, double fontSize) => font switch {
        PdfStandardFont.Courier or PdfStandardFont.CourierBold => fontSize * 0.23,
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold => fontSize * 0.22,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold => fontSize * 0.26,
        _ => fontSize * 0.23
    };

    private static bool LooksNumeric(string s) {
        if (string.IsNullOrWhiteSpace(s)) return false;
        s = s.Trim();
        // currency/percent/simple numeric with separators
        // Use fast char overloads where available; fall back to string overloads on older TFMs.
#if NET8_0_OR_GREATER
        if (s.StartsWith('$') || s.EndsWith('%')) return true;
#else
        if (s.StartsWith("$", System.StringComparison.Ordinal) || s.EndsWith("%", System.StringComparison.Ordinal)) return true;
#endif
        int digits = 0;
        foreach (char ch in s) {
            if (char.IsDigit(ch)) digits++;
            else if (ch == ',' || ch == '.' || ch == ' ' || ch == '+' || ch == '-' || ch == '$') continue;
            else return false;
        }
        return digits > 0;
    }
}
