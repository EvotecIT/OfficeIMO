using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

/// <summary>
/// Page extraction from documents whose bookmarks use named destinations and actions, with every dictionary packed in an
/// object stream behind a cross-reference stream, as LaTeX/hyperref and similar producers write them.
/// </summary>
public sealed class PdfPageExtractorOutlineTests {
    [Fact]
    public void ExtractAllPagesKeepsNestedOutlineNamedDestinationsAndLinks() {
        byte[] source = BuildCompressedOutlinePdf();
        PdfDocumentInfo sourceInfo = PdfInspector.Inspect(source);

        using PdfDocument extracted = PdfDocument.Load(source).Pages.Extract("1-3");

        byte[] output = extracted.ToBytes();
        PdfDocumentInfo info = PdfInspector.Inspect(output);
        Assert.Equal(Flatten(sourceInfo.Outlines), Flatten(info.Outlines));
        Assert.Equal(new[] { "chapter.1", "chapter.1", "chapter.3", "part.1" }, info.NamedDestinations.Select(destination => destination.Name).OrderBy(name => name, StringComparer.Ordinal));
        Assert.Equal(3, info.LinkAnnotations.Count);
        Assert.Empty(extracted.Pipeline.Steps.Last().Diagnostics);
        AssertReadable(output, 3, "Page three");
    }

    [Fact]
    public void ExtractSubsetPrunesOutlineItemsForRemovedPagesAndReportsEachChange() {
        byte[] source = BuildCompressedOutlinePdf();

        using PdfDocument extracted = PdfDocument.Load(source).Pages.Extract("1-2");

        byte[] output = extracted.ToBytes();
        PdfDocumentInfo info = PdfInspector.Inspect(output);
        Assert.Equal(new[] {
            "1 Part I -> 1 (open)",
            "2 Chapter 1 -> 1",
            "2 Chapter 2 -> 2",
            "1 Part II -> none (closed)",
            "2 Project site -> none"
        }, Flatten(info.Outlines));
        Assert.Equal(new[] { "chapter.1", "part.1" }, info.NamedDestinations.Select(destination => destination.Name).OrderBy(name => name, StringComparer.Ordinal));
        PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);
        Assert.Equal("chapter.1", link.DestinationName);

        IReadOnlyList<string> report = extracted.Pipeline.Steps.Last().Diagnostics;
        Assert.Equal(4, report.Count);
        Assert.Contains(report, entry => entry.Contains("\"Chapter 3\" was removed", StringComparison.Ordinal));
        Assert.Contains(report, entry => entry.Contains("\"Part II\" was kept without a destination", StringComparison.Ordinal));
        Assert.Contains(report, entry => entry.StartsWith("2 named destinations ", StringComparison.Ordinal));
        Assert.Contains(report, entry => entry.StartsWith("2 link annotations ", StringComparison.Ordinal));
        AssertReadable(output, 2, "Page two");
    }

    [Fact]
    public void SplitReportsNavigationRemovedFromEachPart() {
        byte[] source = BuildCompressedOutlinePdf();

        IReadOnlyList<PdfDocument> parts = PdfDocument.Load(source).Pages.Split("1-2,3");

        Assert.Equal(4, parts[0].Pipeline.Steps.Last().Diagnostics.Count);
        PdfDocumentInfo last = PdfInspector.Inspect(parts[1].ToBytes());
        Assert.Equal(new[] { "1 Part II -> 1 (closed)", "2 Chapter 3 -> 1", "2 Project site -> none" }, Flatten(last.Outlines));
        IReadOnlyList<string> report = parts[1].Pipeline.Steps.Last().Diagnostics;
        Assert.Contains(report, entry => entry.Contains("\"Part I\" was removed", StringComparison.Ordinal));
        Assert.Contains(report, entry => entry.Contains("\"Chapter 1\" was removed", StringComparison.Ordinal));
        Assert.Contains(report, entry => entry.Contains("\"Chapter 2\" was removed", StringComparison.Ordinal));
        foreach (PdfDocument part in parts) part.Dispose();
    }

    [Fact]
    public void ExtractRefusesBookmarkedDocumentReferencingMissingObjectWithTypedReason() {
        // Object 99 does not exist. The catalog-preserving writer would copy it with the bookmark, so extraction is
        // refused while planning instead of failing inside the writer.
        byte[] source = BuildCompressedOutlinePdf(chapterOneExtra: " /SE 99 0 R");
        using PdfDocument document = PdfDocument.Load(source);

        var blocked = Assert.Throws<PdfMutationBlockedException>(() => document.Pages.Extract("1-3"));

        Assert.Contains("FullRewrite.InvalidObjectReferences", blocked.Plan.BlockerCodes);
        Assert.Contains(blocked.Plan.Diagnostics, diagnostic => diagnostic.Contains("PDF object 99 was referenced but not found.", StringComparison.Ordinal));
        Assert.False(document.PlanMutation(PdfMutationOperation.ExtractPages).CanExecute);
    }

    private static void AssertReadable(byte[] output, int pageCount, string text) {
        PdfDocumentPreflight preflight = PdfInspector.Preflight(output);
        Assert.True(preflight.CanRead, string.Join(" ", preflight.Diagnostics));
        Assert.False(preflight.HasRewriteBlocker(PdfRewriteBlockerKind.InvalidObjectReferences));
        using PdfDocument reopened = PdfDocument.Load(output);
        Assert.Equal(pageCount, reopened.Pages.Count);
        Assert.Contains(text, reopened.ExtractText(), StringComparison.Ordinal);
    }

    private static string[] Flatten(IReadOnlyList<PdfOutlineItem> items) {
        var lines = new List<string>();
        Add(items);
        return lines.ToArray();

        void Add(IReadOnlyList<PdfOutlineItem> level) {
            foreach (PdfOutlineItem item in level) {
                string page = item.PageNumber?.ToString(CultureInfo.InvariantCulture) ?? "none";
                string state = item.Children.Count == 0 ? string.Empty : item.IsExpanded ? " (open)" : " (closed)";
                lines.Add(item.Level.ToString(CultureInfo.InvariantCulture) + " " + item.Title + " -> " + page + state);
                Add(item.Children);
            }
        }
    }

    /// <summary>
    /// Three pages with a two-level outline (named, explicit and URI targets), a /Names /Dests tree with Kids and Limits,
    /// and named-destination links. The catalog /Dests name object /chapter.1 goes to page three while the tree's string
    /// (chapter.1) goes to page one. Every dictionary sits in one object stream behind an xref stream.
    /// </summary>
    private static byte[] BuildCompressedOutlinePdf(string chapterOneExtra = "") {
        var packed = new List<(int Number, string Body)> {
            (1, "<< /Type /Catalog /Pages 2 0 R /Outlines 20 0 R /Names << /Dests 30 0 R >> /Dests << /chapter.1 [5 0 R /Fit] >> /PageMode /UseOutlines >>"),
            (2, "<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R] /Count 3 /MediaBox [0 0 612 792] /Resources << /Font << /F1 6 0 R >> >> >>"),
            (3, "<< /Type /Page /Parent 2 0 R /Contents 40 0 R /Annots [7 0 R 8 0 R] >>"),
            (4, "<< /Type /Page /Parent 2 0 R /Contents 41 0 R /Annots [9 0 R] >>"),
            (5, "<< /Type /Page /Parent 2 0 R /Contents 42 0 R >>"),
            (6, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"),
            (7, "<< /Type /Annot /Subtype /Link /Rect [72 600 200 620] /Border [0 0 0] /A << /S /GoTo /D (chapter.3) >> >>"),
            (8, "<< /Type /Annot /Subtype /Link /Rect [72 560 200 580] /Border [0 0 0] /Dest (chapter.1) >>"),
            (9, "<< /Type /Annot /Subtype /Link /Rect [72 560 200 580] /Border [0 0 0] /Dest /chapter.1 >>"),
            (20, "<< /Type /Outlines /First 21 0 R /Last 24 0 R /Count 4 >>"),
            (21, "<< /Title (Part I) /Parent 20 0 R /Next 24 0 R /First 22 0 R /Last 23 0 R /Count 2 /Dest (part.1) >>"),
            (22, "<< /Title (Chapter 1) /Parent 21 0 R /Next 23 0 R /A << /S /GoTo /D (chapter.1) >>" + chapterOneExtra + " >>"),
            (23, "<< /Title (Chapter 2) /Parent 21 0 R /Prev 22 0 R /A << /S /GoTo /D [4 0 R /Fit] >> >>"),
            (24, "<< /Title (Part II) /Parent 20 0 R /Prev 21 0 R /First 25 0 R /Last 26 0 R /Count -2 /Dest [5 0 R /XYZ 0 792 0] >>"),
            (25, "<< /Title (Chapter 3) /Parent 24 0 R /Next 26 0 R /A << /S /GoTo /D (chapter.3) >> >>"),
            (26, "<< /Title (Project site) /Parent 24 0 R /Prev 25 0 R /A << /S /URI /URI (https://example.com/) >> >>"),
            (30, "<< /Kids [31 0 R 32 0 R] >>"),
            (31, "<< /Limits [(chapter.1) (chapter.3)] /Names [(chapter.1) [3 0 R /Fit] (chapter.3) [5 0 R /Fit]] >>"),
            (32, "<< /Limits [(part.1) (part.1)] /Names [(part.1) [3 0 R /XYZ 0 792 0]] >>")
        };
        string[] pageText = { "Page one", "Page two", "Page three" };
        const int objectStream = 10;
        const int xrefStream = 50;

        using var output = new MemoryStream();
        var offsets = new Dictionary<int, long>();
        Write(output, "%PDF-1.5\n%âãÏÓ\n");
        for (int index = 0; index < pageText.Length; index++) {
            WriteStream(output, offsets, 40 + index, string.Empty,
                Encoding.ASCII.GetBytes("BT /F1 18 Tf 72 720 Td (" + pageText[index] + ") Tj ET"));
        }

        var header = new StringBuilder();
        var body = new StringBuilder();
        foreach (var (number, text) in packed) {
            header.Append(number.ToString(CultureInfo.InvariantCulture)).Append(' ')
                .Append(body.Length.ToString(CultureInfo.InvariantCulture)).Append(' ');
            body.Append(text).Append('\n');
        }
        byte[] headerBytes = Encoding.ASCII.GetBytes(header.ToString());
        WriteStream(output, offsets, objectStream,
            "/Type /ObjStm /N " + packed.Count.ToString(CultureInfo.InvariantCulture) + " /First " + headerBytes.Length.ToString(CultureInfo.InvariantCulture),
            headerBytes.Concat(Encoding.ASCII.GetBytes(body.ToString())).ToArray());

        long xrefOffset = output.Position;
        var entries = new List<byte>();
        for (int number = 0; number <= xrefStream; number++) {
            int packedIndex = packed.FindIndex(item => item.Number == number);
            if (packedIndex >= 0) AddEntry(2, objectStream, packedIndex);
            else if (number == xrefStream) AddEntry(1, (int)xrefOffset, 0);
            else if (offsets.TryGetValue(number, out long offset)) AddEntry(1, (int)offset, 0);
            else AddEntry(0, 0, number == 0 ? 65535 : 0);
        }
        Write(output, xrefStream.ToString(CultureInfo.InvariantCulture) + " 0 obj\n<< /Type /XRef /Size " +
            (xrefStream + 1).ToString(CultureInfo.InvariantCulture) + " /W [1 4 2] /Root 1 0 R /Length " +
            entries.Count.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(entries.ToArray(), 0, entries.Count);
        Write(output, "\nendstream\nendobj\nstartxref\n" + xrefOffset.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
        return output.ToArray();

        void AddEntry(int type, int field1, int field2) {
            entries.Add((byte)type);
            for (int shift = 24; shift >= 0; shift -= 8) entries.Add((byte)(field1 >> shift));
            entries.Add((byte)(field2 >> 8));
            entries.Add((byte)field2);
        }
    }

    private static void WriteStream(MemoryStream output, Dictionary<int, long> offsets, int number, string dictionary, byte[] data) {
        offsets[number] = output.Position;
        Write(output, number.ToString(CultureInfo.InvariantCulture) + " 0 obj\n<< " + dictionary + " /Length " +
            data.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(data, 0, data.Length);
        Write(output, "\nendstream\nendobj\n");
    }

    private static void Write(MemoryStream output, string text) {
        byte[] bytes = Encoding.GetEncoding("ISO-8859-1").GetBytes(text);
        output.Write(bytes, 0, bytes.Length);
    }
}
