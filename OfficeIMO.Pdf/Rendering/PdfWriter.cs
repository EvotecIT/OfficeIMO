using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    public static byte[] Write(PdfDoc doc, IEnumerable<IPdfBlock> blocks, PdfOptions opts, string? title, string? author, string? subject, string? keywords) {
        // Layout blocks into pages and create per-page content streams.
        var layout = LayoutBlocks(blocks, opts);

        // Build PDF objects as byte arrays, then assemble with xref.
        var objects = new List<byte[]>();

        // Reserve IDs (1-based). We'll assign as we add to `objects`.
        int infoId = 0, catalogId = 0, pagesId = 0;
        var pageIds = new List<int>();
        var contentIds = new List<int>();

        // Collect fonts used across pages
        var fontObjectIds = new Dictionary<PdfStandardFont, int>();
        int EnsureFont(PdfStandardFont font) {
            if (!fontObjectIds.TryGetValue(font, out int id)) {
                id = AddObject(objects, "<< /Type /Font /Subtype /Type1 /BaseFont /" + font.ToBaseFontName() + " /Encoding /WinAnsiEncoding >>\n");
                fontObjectIds[font] = id;
            }
            return id;
        }

        // Create content streams and page objects
        int totalPages = layout.Pages.Count;
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++) {
            var page = layout.Pages[pageIndex];
            // Make a resources dict that references the fonts we declared
            var pageOpts = page.Options ?? opts;
            var normalFont = ChooseNormal(pageOpts.DefaultFont);
            int normalFontId = EnsureFont(normalFont);
            var fontParts = new List<string> { $"/F1 {normalFontId} 0 R" };
            if (page.UsedBold) {
                var boldFont = ChooseBold(normalFont);
                fontParts.Add($"/F2 {EnsureFont(boldFont)} 0 R");
            }
            if (page.UsedItalic) {
                var italicFont = ChooseItalic(normalFont);
                fontParts.Add($"/F3 {EnsureFont(italicFont)} 0 R");
            }
            if (page.UsedBoldItalic) {
                var biFont = ChooseBoldItalic(normalFont);
                fontParts.Add($"/F4 {EnsureFont(biFont)} 0 R");
            }
            string fontDict = $"<< {string.Join(" ", fontParts)} >>";

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
            if (pageOpts.ShowPageNumbers) {
                string footer = BuildFooter(pageOpts, pageIndex + 1, totalPages);
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
            string annotsPart = pageAnnotIds.Count > 0 ? " /Annots [ " + string.Join(" ", pageAnnotIds.Select(id => id.ToString(CultureInfo.InvariantCulture) + " 0 R")) + " ]" : string.Empty;

            // Page object
            int pageId = AddObject(objects,
                "<< /Type /Page /Parent 0 0 R /MediaBox [0 0 " + F0(pageOpts.PageWidth) + " " + F0(pageOpts.PageHeight) +
                "] /Resources << /Font " + fontDict + (xobjects.Count > 0 ? " /XObject << " + string.Join(" ", xobjects.Select(x => x.Name + " " + x.Id.ToString(CultureInfo.InvariantCulture) + " 0 R")) + " >>" : string.Empty) +
                " >> /Contents " + contentId.ToString(CultureInfo.InvariantCulture) + " 0 R" + annotsPart + " >>\n");
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
}

