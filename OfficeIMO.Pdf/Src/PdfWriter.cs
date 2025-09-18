using System.Text;
using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfWriter {
    private sealed class FontRef {
        public string Name { get; }
        public PdfStandardFont Font { get; }
        public int ObjectId { get; }
        public FontRef(string name, PdfStandardFont font, int objectId) { Name = name; Font = font; ObjectId = objectId; }
    }

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
        fontNormalId = AddObject(objects, $"<< /Type /Font /Subtype /Type1 /BaseFont /{baseFont.ToBaseFontName()} >>\n");
        fonts.Add(new FontRef("F1", baseFont, fontNormalId));
        if (needsBold) {
            var boldFont = ChooseBold(baseFont);
            fontBoldId = AddObject(objects, $"<< /Type /Font /Subtype /Type1 /BaseFont /{boldFont.ToBaseFontName()} >>\n");
            fonts.Add(new FontRef("F2", boldFont, fontBoldId));
        }

        // Create content streams and page objects
        foreach (var page in layout.Pages) {
            // Make a resources dict that references the fonts we declared
            string fontDict = needsBold
                ? $"<< /F1 {fontNormalId} 0 R /F2 {fontBoldId} 0 R >>"
                : $"<< /F1 {fontNormalId} 0 R >>";

            // Content stream
            byte[] content = Encoding.ASCII.GetBytes(page.Content);
            int contentId = AddObject(objects, $"<< /Length {content.Length} >>\nstream\n");
            // Append raw content bytes + endstream/endobj
            // We'll append extra to the last object content after we compute indices; here we simply merge bytes.
            // For simplicity, rebuild the last object with full content now.
            objects[contentId - 1] = Merge(
                Encoding.ASCII.GetBytes($"{contentId} 0 obj\n<< /Length {content.Length} >>\nstream\n"),
                content,
                Encoding.ASCII.GetBytes("\nendstream\nendobj\n")
            );
            contentIds.Add(contentId);

            // Page object
            int pageId = AddObject(objects,
                $"<< /Type /Page /Parent 0 0 R /MediaBox [0 0 {F0(opts.PageWidth)} {F0(opts.PageHeight)}] /Resources << /Font {fontDict} >> /Contents {contentId} 0 R >>\n");
            pageIds.Add(pageId);
        }

        // Pages tree
        string kids = string.Join(" ", pageIds.Select(id => $"{id} 0 R"));
        pagesId = AddObject(objects, $"<< /Type /Pages /Count {pageIds.Count} /Kids [ {kids} ] >>\n");

        // Fix Parent references in each page now that we know pagesId.
        for (int i = 0; i < pageIds.Count; i++) {
            int pageId = pageIds[i];
            string original = Encoding.ASCII.GetString(objects[pageId - 1]);
            string fixedObj = original.Replace("/Parent 0 0 R", $"/Parent {pagesId} 0 R");
            objects[pageId - 1] = Encoding.ASCII.GetBytes(fixedObj);
        }

        // Catalog
        catalogId = AddObject(objects, $"<< /Type /Catalog /Pages {pagesId} 0 R >>\n");

        // Info (metadata)
        var info = new StringBuilder("<< ");
        if (!string.IsNullOrEmpty(title)) info.Append($"/Title {PdfString(title!)} ");
        if (!string.IsNullOrEmpty(author)) info.Append($"/Author {PdfString(author!)} ");
        if (!string.IsNullOrEmpty(subject)) info.Append($"/Subject {PdfString(subject!)} ");
        if (!string.IsNullOrEmpty(keywords)) info.Append($"/Keywords {PdfString(keywords!)} ");
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
        sw.WriteLine($"0 {objects.Count + 1}");
        sw.WriteLine("0000000000 65535 f ");
        for (int i = 1; i <= objects.Count; i++) {
            sw.WriteLine($"{offsets[i]:0000000000} 00000 n ");
        }
        sw.WriteLine("trailer");
        sw.WriteLine($"<< /Size {objects.Count + 1} /Root {catalogId} 0 R /Info {infoId} 0 R >>");
        sw.WriteLine("startxref");
        sw.WriteLine(xrefPos.ToString());
        sw.WriteLine("%%EOF");
        sw.Flush();

        return ms.ToArray();
    }

    private static int AddObject(List<byte[]> list, string body) {
        int id = list.Count + 1;
        var bytes = Encoding.ASCII.GetBytes($"{id} 0 obj\n{body}endobj\n");
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
        // Literal string in parentheses with escapes
        var esc = s.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)").Replace("\r", "\\r").Replace("\n", "\\n");
        return $"({esc})";
    }

    private sealed class LayoutResult {
        public List<Page> Pages { get; } = new();
        public bool UsedBold { get; set; }
        public sealed class Page { public string Content { get; set; } = string.Empty; }
    }

    private static LayoutResult LayoutBlocks(IEnumerable<IPdfBlock> blocks, PdfOptions opts) {
        double width = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        double yStart = opts.PageHeight - opts.MarginTop;
        double y = yStart;

        var sb = new StringBuilder();
        var pages = new List<string>();
        bool usedBold = false;

        void FlushPage() { pages.Add(sb.ToString()); sb.Clear(); }
        void EnsurePageIfNeeded() { if (sb.Length == 0 && pages.Count == 0) { /* first page implicit */ } }
        void NewPage() { FlushPage(); y = yStart; }

        // Helper to write a text block at (x, y) with leading and multiple lines
        void WriteLines(string fontRes, double fontSize, double lineHeight, double x, double startY, IReadOnlyList<string> lines) {
            // Begin text object once per block
            sb.Append("BT\n");
            sb.Append($"/{fontRes} {F(fontSize)} Tf\n");
            sb.Append($"{F(lineHeight)} TL\n");
            sb.Append($"1 0 0 1 {F(opts.MarginLeft)} {F(startY)} Tm\n");
            // First line
            for (int i = 0; i < lines.Count; i++) {
                if (i == 0) {
                    sb.Append('(').Append(EscapeText(lines[i])).Append(") Tj\n");
                } else {
                    sb.Append("T*\n");
                    sb.Append('(').Append(EscapeText(lines[i])).Append(") Tj\n");
                }
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
                WriteLines("F2", size, leading, opts.MarginLeft, y, lines);
                y -= needed;
            }
            else if (block is ParagraphBlock pb) {
                double size = opts.DefaultFontSize;
                double leading = size * 1.4;
                var lines = WrapMonospace(pb.Text, width, size, glyphWidthEm);
                double needed = lines.Count * leading + leading * 0.3;
                if (y - needed < opts.MarginBottom) { NewPage(); }
                WriteLines("F1", size, leading, opts.MarginLeft, y, lines);
                y -= needed;
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

    private static string EscapeText(string s) => s
        .Replace("\\", "\\\\")
        .Replace("(", "\\(")
        .Replace(")", "\\)")
        .Replace("\r", "\\r")
        .Replace("\n", "\\n");

    private static List<string> WrapMonospace(string text, double widthPts, double fontSize, double glyphWidthEm) {
        // Estimated chars per line for monospaced font
        double glyphWidth = fontSize * glyphWidthEm;
        int maxChars = Math.Max(8, (int)Math.Floor(widthPts / glyphWidth));
        var words = text.Replace("\r", "").Split(new[] { ' ', '\n', '\t' }, StringSplitOptions.None);
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
}
