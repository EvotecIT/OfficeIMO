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

        // Collect required fonts across all pages (basic: normal + optional bold/italic/bold-italic)
        bool needsBold = layout.UsedBold;
        bool needsItalic = layout.UsedItalic;
        bool needsBoldItalic = layout.UsedBoldItalic;
        var fonts = new List<FontRef>();
        int fontNormalId = 0;
        int fontBoldId = 0;
        int fontItalicId = 0;
        int fontBoldItalicId = 0;

        // Add font objects
        var baseFont = ChooseNormal(opts.DefaultFont);
        fontNormalId = AddObject(objects, "<< /Type /Font /Subtype /Type1 /BaseFont /" + baseFont.ToBaseFontName() + " /Encoding /WinAnsiEncoding >>\n");
        fonts.Add(new FontRef("F1", baseFont, fontNormalId));
        if (needsBold) {
            var boldFont = ChooseBold(baseFont);
            fontBoldId = AddObject(objects, "<< /Type /Font /Subtype /Type1 /BaseFont /" + boldFont.ToBaseFontName() + " /Encoding /WinAnsiEncoding >>\n");
            fonts.Add(new FontRef("F2", boldFont, fontBoldId));
        }
        if (needsItalic) {
            var italicFont = ChooseItalic(baseFont);
            fontItalicId = AddObject(objects, "<< /Type /Font /Subtype /Type1 /BaseFont /" + italicFont.ToBaseFontName() + " /Encoding /WinAnsiEncoding >>\n");
            fonts.Add(new FontRef("F3", italicFont, fontItalicId));
        }
        if (needsBoldItalic) {
            var biFont = ChooseBoldItalic(baseFont);
            fontBoldItalicId = AddObject(objects, "<< /Type /Font /Subtype /Type1 /BaseFont /" + biFont.ToBaseFontName() + " /Encoding /WinAnsiEncoding >>\n");
            fonts.Add(new FontRef("F4", biFont, fontBoldItalicId));
        }

        // Create content streams and page objects
        int totalPages = layout.Pages.Count;
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++) {
            var page = layout.Pages[pageIndex];
            // Make a resources dict that references the fonts we declared
            string fontDict;
            if (needsBold || needsItalic || needsBoldItalic) {
                var parts = new List<string> { $"/F1 {fontNormalId} 0 R" };
                if (needsBold) parts.Add($"/F2 {fontBoldId} 0 R");
                if (needsItalic) parts.Add($"/F3 {fontItalicId} 0 R");
                if (needsBoldItalic) parts.Add($"/F4 {fontBoldItalicId} 0 R");
                fontDict = $"<< {string.Join(" ", parts)} >>";
            } else {
                fontDict = $"<< /F1 {fontNormalId} 0 R >>";
            }

            // Content stream (append image draw commands at end)
            string contentStr = page.Content;
            var xobjects = new List<(string Name, int Id)>();
            if (page.Images.Count > 0) {
                for (int i = 0; i < page.Images.Count; i++) {
                    var img = page.Images[i];
                    string name = "/Im" + (i + 1).ToString(CultureInfo.InvariantCulture);
                    // Add image object (JPEG assumed: /Filter /DCTDecode)
                    int imgLen = img.Data.Length;
                    int imgId = AddObject(objects, "<< /Type /XObject /Subtype /Image /Width " + F0(img.W) + " /Height " + F0(img.H) + " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " + imgLen.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
                    objects[imgId - 1] = Merge(
                        Encoding.ASCII.GetBytes(imgId.ToString(CultureInfo.InvariantCulture) + " 0 obj\n<< /Type /XObject /Subtype /Image /Width " + F0(img.W) + " /Height " + F0(img.H) + " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " + imgLen.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n"),
                        img.Data,
                        Encoding.ASCII.GetBytes("\nendstream\nendobj\n")
                    );
                    img.ObjectId = imgId;
                    img.Name = name;
                    xobjects.Add((name, imgId));
                }
                // Append draw commands
                var sbImgs = new StringBuilder();
                foreach (var img in page.Images) {
                    sbImgs.Append("q ")
                          .Append(F(img.W)).Append(' ').Append("0 0 ")
                          .Append(F(img.H)).Append(' ').Append(F(img.X)).Append(' ').Append(F(img.Y)).Append(" cm ")
                          .Append(img.Name).Append(" Do Q\n");
                }
                contentStr += sbImgs.ToString();
            }
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

            // Annotations (link URIs)
            var pageAnnotIds = new List<int>();
            if (page.Annotations.Count > 0) {
                foreach (var a in page.Annotations) {
                    string annot = "<< /Type /Annot /Subtype /Link /Border [0 0 0] /Rect [" +
                        F(a.X1) + " " + F(a.Y1) + " " + F(a.X2) + " " + F(a.Y2) +
                        "] /A << /S /URI /URI " + PdfString(a.Uri) + " >> >>\n";
                    int annId = AddObject(objects, annot);
                    pageAnnotIds.Add(annId);
                }
            }

            // Page object
            string annotsPart = pageAnnotIds.Count > 0 ? " /Annots [ " + string.Join(" ", pageAnnotIds.Select(id => id.ToString(CultureInfo.InvariantCulture) + " 0 R")) + " ]" : string.Empty;
            string xobjPart = xobjects.Count > 0 ? " /XObject << " + string.Join(" ", xobjects.Select(x => x.Name + " " + x.Id.ToString(CultureInfo.InvariantCulture) + " 0 R")) + " >>" : string.Empty;
            int pageId = AddObject(objects,
                "<< /Type /Page /Parent 0 0 R /MediaBox [0 0 " + F0(opts.PageWidth) + " " + F0(opts.PageHeight) + 
                "] /Resources << /Font " + fontDict + (xobjPart.Length > 0 ? " /XObject " + xobjPart : string.Empty) + " >> /Contents " + contentId.ToString(CultureInfo.InvariantCulture) + " 0 R" + annotsPart + " >>\n");
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
        public bool UsedItalic { get; set; }
        public bool UsedBoldItalic { get; set; }
        public sealed class Page {
            public string Content { get; set; } = string.Empty;
            public List<LinkAnnotation> Annotations { get; } = new();
            public List<PageImage> Images { get; } = new();
        }
    }

    private sealed class LinkAnnotation {
        public double X1 { get; init; }
        public double Y1 { get; init; }
        public double X2 { get; init; }
        public double Y2 { get; init; }
        public string Uri { get; init; } = string.Empty;
    }

    private sealed class PageImage {
        public byte[] Data { get; init; } = System.Array.Empty<byte>();
        public double X { get; init; }
        public double Y { get; init; }
        public double W { get; init; }
        public double H { get; init; }
        public string Name { get; set; } = string.Empty;
        public int ObjectId { get; set; }
    }

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
        var pages = new List<LayoutResult.Page>();
        var currentPage = new LayoutResult.Page();
        bool usedBold = false;
        bool usedItalic = false;
        bool usedBoldItalic = false;

        void FlushPage() { currentPage.Content = sb.ToString(); pages.Add(currentPage); currentPage = new LayoutResult.Page(); sb.Clear(); }
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
            sb.Append("1 0 0 1 ").Append(F(x)).Append(' ').Append(F(yStart)).Append(" Tm\n");
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
                // If single-line heading and link provided, record annotation rect
                bool singleLine = lines.Count == 1 && !string.IsNullOrEmpty(hb.LinkUri);
                double widthContent = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
                double em = GlyphWidthEmFor(ChooseBold(ChooseNormal(opts.DefaultFont)));
                double lineWidth = lines[0].Length * size * em;
                double dx = 0;
                if (hb.Align == PdfAlign.Center) dx = Math.Max(0, (widthContent - lineWidth) / 2);
                else if (hb.Align == PdfAlign.Right) dx = Math.Max(0, widthContent - lineWidth);
                if (singleLine) {
                    double baseline = y;
                    var baseFont = ChooseBold(ChooseNormal(opts.DefaultFont));
                    double asc = GetAscender(baseFont, size);
                    double desc = GetDescender(baseFont, size);
                    double x1 = opts.MarginLeft + dx;
                    double x2 = x1 + lineWidth;
                    double y1 = baseline - desc;
                    double y2 = baseline + asc;
                    currentPage.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = hb.LinkUri! });
                }
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
            else if (block is RichParagraphBlock rpb) {
                // Inline styled paragraph: detect style usage for font resource registration
                if (rpb.Runs.Any(run => run.Bold && run.Italic)) usedBoldItalic = true;
                if (rpb.Runs.Any(run => run.Bold)) usedBold = true;
                if (rpb.Runs.Any(run => run.Italic)) usedItalic = true;
                double size = opts.DefaultFontSize;
                double leading = size * 1.4;
                var (lines, lineHeights) = WrapRichRuns(rpb.Runs, width, size, ChooseNormal(opts.DefaultFont));
                double totalNeeded = lineHeights.Sum() + leading * 0.3;
                if (y - totalNeeded < opts.MarginBottom) { NewPage(); }
                WriteRichParagraph(sb, rpb, lines, lineHeights, opts, y, size, leading, currentPage.Annotations);
                y -= totalNeeded;
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
                        // Link for cell?
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
            else if (block is RowBlock rb) {
                // Compute column x positions
                double contentWidth = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
                double[] colXs = new double[rb.Columns.Count];
                double[] colWs = new double[rb.Columns.Count];
                double xAcc = opts.MarginLeft;
                for (int i = 0; i < rb.Columns.Count; i++) {
                    double w = System.Math.Max(0, contentWidth * (rb.Columns[i].WidthPercent / 100.0));
                    colXs[i] = xAcc;
                    colWs[i] = w;
                    xAcc += w;
                }
                // Render each column separately, track max height
                double[] colHeights = new double[rb.Columns.Count];
                for (int ci = 0; ci < rb.Columns.Count; ci++) {
                    double xCol = colXs[ci];
                    double wCol = colWs[ci];
                    double yCol = y;
                    foreach (var cb in rb.Columns[ci].Blocks) {
                        if (cb is HeadingBlock hb2) {
                            double size = hb2.Level switch { 1 => 24, 2 => 18, 3 => 14, _ => 12 };
                            double leading = size * 1.25;
                            var lines = WrapMonospace(hb2.Text, wCol, size, GlyphWidthEmFor(ChooseBold(ChooseNormal(opts.DefaultFont))));
                            double needed = lines.Count * leading + leading * 0.25;
                            WriteLines("F2", size, leading, xCol, yCol, lines, hb2.Align, hb2.Color, applyBaselineTweak: false);
                            yCol -= needed;
                        } else if (cb is RichParagraphBlock rpb2) {
                            double size = opts.DefaultFontSize;
                            double leading = size * 1.4;
                            var (lines, lineHeights) = WrapRichRuns(rpb2.Runs, wCol, size, ChooseNormal(opts.DefaultFont));
                            double needed = lineHeights.Sum() + leading * 0.3;
                            WriteRichParagraph(sb, rpb2, lines, lineHeights, opts, yCol, size, leading, currentPage.Annotations, xCol);
                            yCol -= needed;
                        } else if (cb is HorizontalRuleBlock hr2) {
                            yCol -= hr2.SpacingBefore;
                            double x1 = xCol;
                            double x2 = xCol + wCol;
                            double yLine = yCol - hr2.Thickness * 0.5;
                            DrawHLine(sb, hr2.Color, System.Math.Max(0.2, hr2.Thickness), x1, x2, yLine);
                            yCol -= hr2.SpacingAfter;
                        } else if (cb is ImageBlock ib2) {
                            double xImg = xCol;
                            if (ib2.Align == PdfAlign.Center) xImg = xCol + System.Math.Max(0, (wCol - ib2.Width) / 2);
                            else if (ib2.Align == PdfAlign.Right) xImg = xCol + System.Math.Max(0, wCol - ib2.Width);
                            currentPage.Images.Add(new PageImage { Data = ib2.Data, X = xImg, Y = yCol - ib2.Height, W = ib2.Width, H = ib2.Height });
                            yCol -= ib2.Height;
                        }
                    }
                    colHeights[ci] = y - yCol;
                }
                double rowHeight = colHeights.Length > 0 ? colHeights.Max() : 0;
                if (y - rowHeight < opts.MarginBottom) { NewPage(); }
                y -= rowHeight;
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
                var (lines, lineHeights) = WrapRichRuns(ppb.Runs, width, size, ChooseNormal(opts.DefaultFont));
                double textHeight = lineHeights.Sum();
                double panelTop = y;
                // Compute panel width including horizontal padding
                double contentWidth = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
                double innerWidth = contentWidth;
                if (ppb.Style.MaxWidth.HasValue) innerWidth = System.Math.Min(contentWidth, ppb.Style.MaxWidth.Value);
                double xLeft = opts.MarginLeft;
                if (ppb.Style.Align == PdfAlign.Center) xLeft = opts.MarginLeft + System.Math.Max(0, (contentWidth - innerWidth) / 2);
                else if (ppb.Style.Align == PdfAlign.Right) xLeft = opts.MarginLeft + System.Math.Max(0, contentWidth - innerWidth);
                double panelBottom = y - (ppb.Style.PaddingY + textHeight + ppb.Style.PaddingY);
                if (panelBottom < opts.MarginBottom) { NewPage(); panelTop = y; panelBottom = y - (ppb.Style.PaddingY + textHeight + ppb.Style.PaddingY); }
                double panelWidth = innerWidth;
                if (ppb.Style.Background.HasValue) {
                    DrawRowFill(sb, ppb.Style.Background.Value, xLeft, panelBottom, panelWidth, panelTop - panelBottom);
                }
                if (ppb.Style.BorderColor.HasValue && ppb.Style.BorderWidth > 0) {
                    DrawRowRect(sb, ppb.Style.BorderColor.Value, ppb.Style.BorderWidth, xLeft, panelBottom, panelWidth, panelTop - panelBottom);
                }
                // Write text starting at y = panelTop - paddingY
                WriteRichParagraph(sb, new RichParagraphBlock(ppb.Runs, ppb.Align, ppb.DefaultColor), lines, lineHeights, opts, panelTop - ppb.Style.PaddingY, size, leading, currentPage.Annotations, xLeft + ppb.Style.PaddingX);
                y = panelBottom;
            }
            else {
                // Unknown future block types – skip for now
            }

            // If we've run out of space exactly, force a new page for the next block
            if (y <= opts.MarginBottom + 4) NewPage();
        }

        // Close last page
        FlushPage();

        var result = new LayoutResult { UsedBold = usedBold, UsedItalic = usedItalic, UsedBoldItalic = usedBoldItalic };
        foreach (var p in pages) result.Pages.Add(p);
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
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaOblique or PdfStandardFont.HelveticaBold => PdfStandardFont.HelveticaBold,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBold => PdfStandardFont.TimesBold,
        PdfStandardFont.Courier or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBold => PdfStandardFont.CourierBold,
        _ => PdfStandardFont.CourierBold
    };

    private static PdfStandardFont ChooseItalic(PdfStandardFont normal) => normal switch {
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold => PdfStandardFont.HelveticaOblique,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold => PdfStandardFont.TimesItalic,
        PdfStandardFont.Courier or PdfStandardFont.CourierBold => PdfStandardFont.CourierOblique,
        _ => PdfStandardFont.HelveticaOblique
    };

    private static PdfStandardFont ChooseBoldItalic(PdfStandardFont normal) => normal switch {
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaOblique or PdfStandardFont.HelveticaBold => PdfStandardFont.HelveticaBoldOblique,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBold => PdfStandardFont.TimesBoldItalic,
        PdfStandardFont.Courier or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBold => PdfStandardFont.CourierBoldOblique,
        _ => PdfStandardFont.HelveticaBoldOblique
    };

    private static double GlyphWidthEmFor(PdfStandardFont font) => font switch {
        PdfStandardFont.Courier or PdfStandardFont.CourierBold or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBoldOblique => 0.6,
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold or PdfStandardFont.HelveticaOblique or PdfStandardFont.HelveticaBoldOblique => 0.55,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBoldItalic => 0.5,
        _ => 0.6
    };

    // Approximate descender height (distance from baseline down to glyph bottom) in points.
    private static double GetDescender(PdfStandardFont font, double fontSize) => font switch {
        PdfStandardFont.Courier or PdfStandardFont.CourierBold => fontSize * 0.23,
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold => fontSize * 0.22,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold => fontSize * 0.26,
        _ => fontSize * 0.23
    };

    // Approximate ascender height (distance from baseline up to glyph top) in points.
    private static double GetAscender(PdfStandardFont font, double fontSize) => font switch {
        PdfStandardFont.Courier or PdfStandardFont.CourierBold => fontSize * 0.72,
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold => fontSize * 0.74,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold => fontSize * 0.72,
        _ => fontSize * 0.72
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

    // Rich paragraph layout
    private sealed record RichSeg(string Text, bool Bold, bool Italic, bool Underline, bool Strike, PdfColor? Color, string? Uri);

    private static (List<List<RichSeg>> Lines, List<double> LineHeights) WrapRichRuns(IEnumerable<TextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont) {
        // Monospace-ish width estimate per font family; bold uses same em width here.
        double em = GlyphWidthEmFor(baseFont);
        double spaceW = fontSize * em;
        int maxChars = Math.Max(1, (int)System.Math.Floor(maxWidthPts / (fontSize * em)));
        var lines = new List<List<RichSeg>> { new() };
        var heights = new List<double>();
        double lineWidth = 0;
        foreach (var run in runs) {
            string text = run.Text ?? string.Empty;
            bool bold = run.Bold;
            bool underline = run.Underline;
            bool strike = run.Strike;
            bool italic = run.Italic;
            var color = run.Color;
            string? uri = run.LinkUri;
            int idx = 0;
            while (idx < text.Length) {
                // find next whitespace or newline as a break opportunity
                int nextWs = text.IndexOfAny(new[] { ' ', '\n' }, idx);
                bool hadNewline = false;
                string token;
                if (nextWs == -1) { token = text.Substring(idx); idx = text.Length; }
                else {
                    token = text.Substring(idx, nextWs - idx);
                    hadNewline = text[nextWs] == '\n';
                    idx = nextWs + 1;
                }
                double tokenW = token.Length * fontSize * em;
                var lastLine = lines[lines.Count - 1];
                double needed = (lastLine.Count == 0 ? tokenW : spaceW + tokenW);

                // Hard break overly long tokens
                if (tokenW > maxWidthPts) {
                    if (lastLine.Count > 0) { heights.Add(fontSize * 1.4); lines.Add(new()); lineWidth = 0; lastLine = lines[lines.Count - 1]; }
                    int pos = 0;
                    while (pos < token.Length) {
                        int take = System.Math.Min(maxChars, token.Length - pos);
                        string chunk = token.Substring(pos, take);
                        lastLine.Add(new RichSeg(chunk, bold, italic, underline, strike, color, uri));
                        pos += take;
                        if (pos < token.Length) { heights.Add(fontSize * 1.4); lines.Add(new()); lineWidth = 0; lastLine = lines[lines.Count - 1]; }
                    }
                    if (hadNewline) { heights.Add(fontSize * 1.4); lines.Add(new()); lineWidth = 0; }
                    continue;
                }
                if (lineWidth + needed > maxWidthPts && lastLine.Count > 0) {
                    // break line
                    heights.Add(fontSize * 1.4);
                    lines.Add(new());
                    lineWidth = 0;
                    needed = tokenW;
                }
                if (token.Length > 0) {
                    if (lineWidth > 0) lineWidth += spaceW;
                    lines[lines.Count - 1].Add(new RichSeg(token, bold, italic, underline, strike, color, uri));
                    lineWidth += tokenW;
                }
                if (hadNewline) {
                    heights.Add(fontSize * 1.4);
                    lines.Add(new());
                    lineWidth = 0;
                }
            }
        }
        if (lines.Count > 0 && lines[lines.Count - 1].Count == 0) { lines.RemoveAt(lines.Count - 1); }
        if (heights.Count < lines.Count) heights.Add(fontSize * 1.4);
        return (lines, heights);
    }

    private static void WriteRichParagraph(StringBuilder sb, RichParagraphBlock block, List<List<RichSeg>> lines, List<double> lineHeights, PdfOptions opts, double startY, double fontSize, double defaultLeading, List<LinkAnnotation> annots, double? xOverride = null) {
        double widthContent = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        // Precompute run widths using monospace em for default font
        double em = GlyphWidthEmFor(ChooseNormal(opts.DefaultFont));
        double spaceW = fontSize * em;
        // Underlines to draw after text
        var underlines = new List<(double X1, double X2, double Y, PdfColor Color)>();
        var strikes = new List<(double X1, double X2, double Y, PdfColor Color)>();

        sb.Append("BT\n");
        sb.Append(F(defaultLeading)).Append(" TL\n");
        double xOrigin = xOverride ?? opts.MarginLeft;
        sb.Append("1 0 0 1 ").Append(F(xOrigin)).Append(' ').Append(F(startY)).Append(" Tm\n");

        for (int li = 0; li < lines.Count; li++) {
            if (li != 0) sb.Append("T*\n");
            var segs = lines[li];
            // measure line width
            double lineW = 0;
            foreach (var s in segs) lineW += (s.Text.Length * fontSize * em);
            // add spaces between tokens if present
            if (segs.Count > 1) lineW += (segs.Count - 1) * spaceW;

            double dx = 0;
            if (block.Align == PdfAlign.Center) dx = Math.Max(0, (widthContent - lineW) / 2);
            else if (block.Align == PdfAlign.Right) dx = Math.Max(0, widthContent - lineW);
            if (dx != 0) sb.Append(F(dx)).Append(" 0 Td\n");

            double xCursor = dx;
            for (int si = 0; si < segs.Count; si++) {
                var s = segs[si];
                // font
                string fontRes = (s.Bold && s.Italic) ? "F4" : s.Bold ? "F2" : s.Italic ? "F3" : "F1";
                sb.Append('/').Append(fontRes).Append(' ').Append(F(fontSize)).Append(" Tf\n");
                // color
                var color = s.Color ?? block.DefaultColor ?? opts.DefaultTextColor;
                if (color.HasValue) sb.Append(SetFillColor(color.Value));
                // text
                sb.Append('<').Append(EncodeWinAnsiHex(s.Text)).Append("> Tj\n");
                double wSeg = s.Text.Length * fontSize * em;
                // explicitly advance by segment width to ensure position progresses across viewers
                sb.Append(F(wSeg)).Append(" 0 Td\n");

                // underline plan
                if (s.Underline) {
                    var ulColor = (s.Color ?? block.DefaultColor ?? opts.DefaultTextColor) ?? PdfColor.Black;
                    double yLine = startY - li * defaultLeading - fontSize * 0.15; // ~ underline offset
                    underlines.Add((xOrigin + xCursor, xOrigin + xCursor + wSeg, yLine, ulColor));
                }
                if (s.Strike) {
                    var stColor = (s.Color ?? block.DefaultColor ?? opts.DefaultTextColor) ?? PdfColor.Black;
                    double yLine = startY - li * defaultLeading + fontSize * 0.32; // ~ strike offset
                    strikes.Add((xOrigin + xCursor, xOrigin + xCursor + wSeg, yLine, stColor));
                }
                // link annotation
                if (!string.IsNullOrEmpty(s.Uri)) {
                    double baseline = startY - li * defaultLeading;
                    var baseFont = ChooseNormal(opts.DefaultFont);
                    var fontForMetrics = (s.Bold && s.Italic) ? ChooseBoldItalic(baseFont) : s.Bold ? ChooseBold(baseFont) : s.Italic ? ChooseItalic(baseFont) : baseFont;
                    double asc = GetAscender(fontForMetrics, fontSize);
                    double desc = GetDescender(fontForMetrics, fontSize);
                    double x1 = xOrigin + xCursor;
                    double x2 = x1 + wSeg;
                    double y1 = baseline - desc;
                    double y2 = baseline + asc;
                    annots.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = s.Uri! });
                }
                // advance x by segment width + inter-token space if the next is not the first
                xCursor += wSeg;
                if (si != segs.Count - 1) { xCursor += spaceW; sb.Append(F(spaceW)).Append(" 0 Td\n"); }
            }
            // Return to start-of-line X for next line
            if (xCursor != 0) sb.Append(F(-xCursor)).Append(" 0 Td\n");
        }
        sb.Append("ET\n");

        // draw underlines
        foreach (var ul in underlines) {
            sb.Append("q\n");
            sb.Append(SetStrokeColor(ul.Color));
            sb.Append("0.5 w\n");
            sb.Append(F(ul.X1)).Append(' ').Append(F(ul.Y)).Append(" m ").Append(F(ul.X2)).Append(' ').Append(F(ul.Y)).Append(" l S\n");
            sb.Append("Q\n");
        }
        foreach (var st in strikes) {
            sb.Append("q\n");
            sb.Append(SetStrokeColor(st.Color));
            sb.Append("0.5 w\n");
            sb.Append(F(st.X1)).Append(' ').Append(F(st.Y)).Append(" m ").Append(F(st.X2)).Append(' ').Append(F(st.Y)).Append(" l S\n");
            sb.Append("Q\n");
        }
    }
}
