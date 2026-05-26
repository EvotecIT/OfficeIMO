using System.IO.Compression;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfExternalDocumentCompatibilityTests {
    [Fact]
    public void ExtractAllText_PreservesSpacingFromExternalTjArrays() {
        byte[] pdf = BuildExternalSinglePagePdf(
            "BT\n/F13 12 Tf\n72 720 Td\n[(External) -360 (PDF) -360 (text)] TJ\nET\n");

        string allText = PdfTextExtractor.ExtractAllText(pdf);
        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));

        Assert.Contains("External PDF text", Normalize(allText), StringComparison.Ordinal);
        Assert.Contains("External PDF text", Normalize(pageText), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_UsesExternalToUnicodeCMap() {
        byte[] pdf = BuildExternalToUnicodePdf();

        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));

        Assert.Contains("Zed", Normalize(pageText), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_UsesExternalToUnicodeBfRangeArrayCMap() {
        byte[] pdf = BuildExternalToUnicodeBfRangeArrayPdf();

        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));

        Assert.Contains("Zfi!", Normalize(pageText), StringComparison.Ordinal);
    }

    [Fact]
    public void SplitPages_ReadsExternalProducerPdfWithInheritedResourcesAndContentArrays() {
        byte[] pdf = BuildExternalTwoPagePdf();

        IReadOnlyList<byte[]> pages = PdfPageExtractor.SplitPages(pdf);

        Assert.Equal(2, pages.Count);
        Assert.Contains("External first page", Normalize(PdfTextExtractor.ExtractAllText(pages[0])), StringComparison.Ordinal);
        Assert.Contains("External second page", Normalize(PdfTextExtractor.ExtractAllText(pages[1])), StringComparison.Ordinal);
    }

    [Fact]
    public void Merge_ReordersExternalProducerPdfPagesAfterSplit() {
        byte[] pdf = BuildExternalTwoPagePdf();
        IReadOnlyList<byte[]> pages = PdfPageExtractor.SplitPages(pdf);

        byte[] merged = PdfMerger.Merge(pages[1], pages[0]);

        PdfDocumentInfo info = PdfInspector.Inspect(merged);
        string text = Normalize(PdfTextExtractor.ExtractAllText(merged));
        int secondPageIndex = text.IndexOf("External second page", StringComparison.Ordinal);
        int firstPageIndex = text.IndexOf("External first page", StringComparison.Ordinal);
        Assert.Equal(2, info.PageCount);
        Assert.NotEqual(-1, secondPageIndex);
        Assert.NotEqual(-1, firstPageIndex);
        Assert.True(secondPageIndex < firstPageIndex, text);
    }

    [Fact]
    public void SplitAndMerge_ReadExternalObjectStreamPageTree() {
        byte[] pdf = BuildExternalObjectStreamPdf(includeAcroForm: false);

        Assert.Contains("Object stream page", Normalize(PdfTextExtractor.ExtractAllText(pdf)), StringComparison.Ordinal);

        IReadOnlyList<byte[]> pages = PdfPageExtractor.SplitPages(pdf);
        Assert.Single(pages);
        Assert.Contains("Object stream page", Normalize(PdfTextExtractor.ExtractAllText(pages[0])), StringComparison.Ordinal);

        byte[] merged = PdfMerger.Merge(pdf, pdf);
        PdfDocumentInfo info = PdfInspector.Inspect(merged);
        string mergedText = Normalize(PdfTextExtractor.ExtractAllText(merged));

        Assert.Equal(2, info.PageCount);
        Assert.Equal(2, CountOccurrences(mergedText, "Object stream page"));
    }

    [Fact]
    public void ExtractText_UsesXrefStreamOffsetsInsteadOfTrailingStaleDuplicateObjects() {
        byte[] pdf = BuildXrefStreamPdfWithTrailingStaleDuplicatePage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Active xref stream page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Stale trailing page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_UsesXrefStreamCompressedObjectEntriesInsteadOfTrailingStaleDuplicates() {
        byte[] pdf = BuildXrefStreamCompressedObjectPdfWithTrailingStaleDuplicatePage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Active compressed xref page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Stale compressed trailing page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_FollowsXrefStreamPrevChainForInheritedObjects() {
        byte[] pdf = BuildIncrementalXrefStreamPdfWithTrailingStaleDuplicatePage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Inherited previous xref page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Stale incremental trailing page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_HonorsXrefStreamFreeEntriesOverTrailingStaleObjects() {
        byte[] pdf = BuildIncrementalXrefStreamPdfWithFreedTrailingStalePage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Active replacement xref stream page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Superseded xref stream page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Stale freed xref stream page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_FollowsClassicXrefPrevChainForInheritedObjects() {
        byte[] pdf = BuildIncrementalClassicXrefPdfWithTrailingStaleDuplicatePage();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Inherited classic xref page", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Stale classic trailing page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Inspect_FollowsClassicTrailerPrevChainForInheritedRoot() {
        byte[] pdf = BuildIncrementalClassicXrefPdfWithInheritedTrailerRoot();

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);

        PdfPageInfo page = Assert.Single(info.Pages);
        Assert.Equal("SinglePage", info.CatalogPageLayout);
        Assert.Equal(200d, page.Width);
        Assert.Equal(200d, page.Height);
    }

    [Fact]
    public void ReadExternalObjectStream_DoesNotOverwriteExplicitIndirectObjects() {
        byte[] pdf = BuildExternalObjectStreamWithExplicitReplacementPdf();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Explicit object wins", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Packed object stream wins", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadExternalObjectStream_LaterObjectStreamReplacesEarlierCompressedObjects() {
        byte[] pdf = BuildExternalObjectStreamWithCompressedReplacementPdf();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Later object stream wins", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Earlier object stream wins", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadExternalObjectStream_LaterObjectStreamReplacesEarlierExplicitObjects() {
        byte[] pdf = BuildExternalObjectStreamReplacingEarlierExplicitPdf();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Later object stream wins", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Earlier explicit object wins", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadExternalObjectStream_IgnoresMalformedLaterObjectHeaders() {
        byte[] pdf = BuildExternalObjectStreamWithMalformedTrailingPageObjectPdf();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Object stream page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadExternalObjectStream_IgnoresMalformedLaterDictionaryObjects() {
        byte[] pdf = BuildExternalObjectStreamWithMalformedTrailingPageDictionaryPdf();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Object stream page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadExternalObjectStream_IgnoresParseFailedLaterDictionaryObjects() {
        byte[] pdf = BuildExternalObjectStreamWithParseFailedTrailingPageDictionaryPdf();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Object stream page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RewritePreflight_DetectsCompressedObjectStreamFormMarkers() {
        byte[] pdf = BuildExternalObjectStreamPdf(includeAcroForm: true);
        string rawPdf = Encoding.ASCII.GetString(pdf);

        Assert.DoesNotContain("AcroForm", rawPdf, StringComparison.Ordinal);
        Assert.Contains("Object stream page", Normalize(PdfTextExtractor.ExtractAllText(pdf)), StringComparison.Ordinal);

        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);
        Assert.True(preflight.CanRead);
        Assert.False(preflight.CanRewrite);
        Assert.True(preflight.HasRewriteBlocker(PdfRewriteBlockerKind.Forms));

        var exception = Assert.Throws<NotSupportedException>(() => PdfPageExtractor.SplitPages(pdf));
        Assert.Contains("PDF form fields are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
    }

    private static byte[] BuildExternalSinglePagePdf(string content) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(content.TrimEnd('\n'));
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents [5 0 R] >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, streamBytes)
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildExternalTwoPagePdf() {
        byte[] firstPartOne = Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(External first) Tj\nET\n");
        byte[] firstPartTwo = Encoding.ASCII.GetBytes(EncodeAsciiHex(Compress(Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n150 720 Td\n( page) Tj\nET\n"))));
        byte[] second = Encoding.ASCII.GetBytes(EncodeAsciiHex(Compress(Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n[(External) -360 (second) -360 (page)] TJ\nET\n"))));

        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 5 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents [6 0 R 7 0 R] >>\nendobj",
            "4 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 8 0 R >>\nendobj",
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(6, firstPartOne),
            BuildStreamObject(7, firstPartTwo, "/Filter [/ASCIIHexDecode /FlateDecode]"),
            BuildStreamObject(8, second, "/Filter [/ASCIIHexDecode /FlateDecode]")
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildExternalToUnicodePdf() {
        byte[] content = Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n<010203> Tj\nET\n");
        const string cmap = "/CIDInit /ProcSet findresource begin\n" +
            "12 dict begin\n" +
            "begincmap\n" +
            "3 beginbfchar\n" +
            "<01> <005A>\n" +
            "<02> <0065>\n" +
            "<03> <0064>\n" +
            "endbfchar\n" +
            "endcmap\n" +
            "CMapName currentdict /CMap defineresource pop\n" +
            "end\n" +
            "end\n";

        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /AAAAAA+Helvetica /Encoding /WinAnsiEncoding /ToUnicode 6 0 R >>\nendobj",
            BuildStreamObject(5, content),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(cmap))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildExternalToUnicodeBfRangeArrayPdf() {
        byte[] content = Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n<010203> Tj\nET\n");
        const string cmap = "/CIDInit /ProcSet findresource begin\n" +
            "12 dict begin\n" +
            "begincmap\n" +
            "1 beginbfrange\n" +
            "<01> <03> [<005A> <0066 0069> <0021>]\n" +
            "endbfrange\n" +
            "endcmap\n" +
            "CMapName currentdict /CMap defineresource pop\n" +
            "end\n" +
            "end\n";

        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /AAAAAA+Helvetica /Encoding /WinAnsiEncoding /ToUnicode 6 0 R >>\nendobj",
            BuildStreamObject(5, content),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(cmap))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildXrefStreamPdfWithTrailingStaleDuplicatePage() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        WriteAscii(stream, "%PDF-1.5\n");
        WriteObject(stream, offsets, 1, "<< /Type /Catalog /Pages 2 0 R >>");
        WriteObject(stream, offsets, 2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 7 0 R >> >> >>");
        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>");
        WriteStreamObject(stream, offsets, 4, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Active xref stream page) Tj\nET\n"));
        WriteObject(stream, offsets, 7, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");

        int xrefObjectNumber = 8;
        offsets[xrefObjectNumber] = (int)stream.Position;
        byte[] xrefEntries = BuildXrefStreamEntries(offsets, xrefObjectNumber);
        WriteAscii(stream, xrefObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 9 /Root 1 0 R /W [1 4 2] /Index [0 9] /Length " +
            xrefEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(xrefEntries, 0, xrefEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");

        WriteAscii(stream, "startxref\n" + offsets[xrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>");
        WriteStreamObject(stream, offsets, 6, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Stale trailing page) Tj\nET\n"));

        return stream.ToArray();
    }

    private static byte[] BuildXrefStreamCompressedObjectPdfWithTrailingStaleDuplicatePage() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        WriteAscii(stream, "%PDF-1.5\n");
        WriteStreamObject(stream, offsets, 4, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Active compressed xref page) Tj\nET\n"));

        var packedObjects = new List<(int ObjectNumber, string Body)> {
            (1, "<< /Type /Catalog /Pages 2 0 R >>"),
            (2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 7 0 R >> >> >>"),
            (3, "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>"),
            (7, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
        };
        WriteRawObject(stream, offsets, 10, BuildObjectStreamObject(10, packedObjects));

        int xrefObjectNumber = 11;
        offsets[xrefObjectNumber] = (int)stream.Position;
        var entries = new Dictionary<int, (int Type, int Field1, int Field2)> {
            [1] = (2, 10, 0),
            [2] = (2, 10, 1),
            [3] = (2, 10, 2),
            [4] = (1, offsets[4], 0),
            [7] = (2, 10, 3),
            [10] = (1, offsets[10], 0),
            [11] = (1, offsets[xrefObjectNumber], 0)
        };
        byte[] xrefEntries = BuildXrefStreamEntries(entries, size: 12);
        WriteAscii(stream, xrefObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 12 /Root 1 0 R /W [1 4 2] /Index [0 12] /Length " +
            xrefEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(xrefEntries, 0, xrefEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");

        WriteAscii(stream, "startxref\n" + offsets[xrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>");
        WriteStreamObject(stream, offsets, 6, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Stale compressed trailing page) Tj\nET\n"));

        return stream.ToArray();
    }

    private static byte[] BuildIncrementalXrefStreamPdfWithTrailingStaleDuplicatePage() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        WriteAscii(stream, "%PDF-1.5\n");
        WriteObject(stream, offsets, 1, "<< /Type /Catalog /Pages 2 0 R >>");
        WriteObject(stream, offsets, 2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 7 0 R >> >> >>");
        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>");
        WriteStreamObject(stream, offsets, 4, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Inherited previous xref page) Tj\nET\n"));
        WriteObject(stream, offsets, 7, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");

        int previousXrefObjectNumber = 8;
        offsets[previousXrefObjectNumber] = (int)stream.Position;
        byte[] previousEntries = BuildXrefStreamEntries(offsets, previousXrefObjectNumber);
        WriteAscii(stream, previousXrefObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 10 /Root 1 0 R /W [1 4 2] /Index [0 9] /Length " +
            previousEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(previousEntries, 0, previousEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");
        WriteAscii(stream, "startxref\n" + offsets[previousXrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        int activeXrefObjectNumber = 9;
        offsets[activeXrefObjectNumber] = (int)stream.Position;
        var activeEntries = new Dictionary<int, (int Type, int Field1, int Field2)> {
            [9] = (1, offsets[activeXrefObjectNumber], 0)
        };
        byte[] xrefEntries = BuildXrefStreamEntries(new[] { 9 }, activeEntries);
        WriteAscii(stream, activeXrefObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 10 /Root 1 0 R /Prev " +
            offsets[previousXrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /W [1 4 2] /Index [9 1] /Length " +
            xrefEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(xrefEntries, 0, xrefEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");
        WriteAscii(stream, "startxref\n" + offsets[activeXrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>");
        WriteStreamObject(stream, offsets, 6, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Stale incremental trailing page) Tj\nET\n"));

        return stream.ToArray();
    }

    private static byte[] BuildIncrementalXrefStreamPdfWithFreedTrailingStalePage() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        WriteAscii(stream, "%PDF-1.5\n");
        WriteObject(stream, offsets, 1, "<< /Type /Catalog /Pages 2 0 R >>");
        WriteObject(stream, offsets, 2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 7 0 R >> >> >>");
        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>");
        WriteStreamObject(stream, offsets, 4, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Superseded xref stream page) Tj\nET\n"));
        WriteObject(stream, offsets, 7, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");

        int previousXrefObjectNumber = 8;
        offsets[previousXrefObjectNumber] = (int)stream.Position;
        byte[] previousEntries = BuildXrefStreamEntries(offsets, previousXrefObjectNumber);
        WriteAscii(stream, previousXrefObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 11 /Root 1 0 R /W [1 4 2] /Index [0 9] /Length " +
            previousEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(previousEntries, 0, previousEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");
        WriteAscii(stream, "startxref\n" + offsets[previousXrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        WriteObject(stream, offsets, 2, "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 7 0 R >> >> >>");
        WriteObject(stream, offsets, 5, "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>");
        WriteStreamObject(stream, offsets, 6, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Active replacement xref stream page) Tj\nET\n"));

        int activeXrefObjectNumber = 9;
        offsets[activeXrefObjectNumber] = (int)stream.Position;
        var activeEntries = new Dictionary<int, (int Type, int Field1, int Field2)> {
            [2] = (1, offsets[2], 0),
            [3] = (0, 0, 65535),
            [5] = (1, offsets[5], 0),
            [6] = (1, offsets[6], 0),
            [9] = (1, offsets[activeXrefObjectNumber], 0)
        };
        byte[] xrefEntries = BuildXrefStreamEntries(new[] { 2, 3, 5, 6, 9 }, activeEntries);
        WriteAscii(stream, activeXrefObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 11 /Root 1 0 R /Prev " +
            offsets[previousXrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /W [1 4 2] /Index [2 2 5 2 9 1] /Length " +
            xrefEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(xrefEntries, 0, xrefEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");
        WriteAscii(stream, "startxref\n" + offsets[activeXrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 10 0 R >>");
        WriteStreamObject(stream, offsets, 10, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Stale freed xref stream page) Tj\nET\n"));

        return stream.ToArray();
    }

    private static byte[] BuildIncrementalClassicXrefPdfWithTrailingStaleDuplicatePage() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        WriteAscii(stream, "%PDF-1.4\n");
        WriteObject(stream, offsets, 1, "<< /Type /Catalog /Pages 2 0 R >>");
        WriteObject(stream, offsets, 2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 7 0 R >> >> >>");
        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>");
        WriteStreamObject(stream, offsets, 4, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Inherited classic xref page) Tj\nET\n"));
        WriteObject(stream, offsets, 7, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");

        int previousXrefOffset = (int)stream.Position;
        var previousEntries = new Dictionary<int, int>(offsets) {
            [0] = 0
        };
        WriteClassicXrefTable(stream, previousEntries, size: 9, rootObjectNumber: 1, previousXrefOffset: null);
        WriteAscii(stream, "startxref\n" + previousXrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        WriteObject(stream, offsets, 8, "<< /Producer (OfficeIMO incremental marker) >>");
        int activeXrefOffset = (int)stream.Position;
        var activeEntries = new Dictionary<int, int> {
            [8] = offsets[8]
        };
        WriteClassicXrefTable(stream, activeEntries, size: 9, rootObjectNumber: 1, previousXrefOffset: previousXrefOffset);
        WriteAscii(stream, "startxref\n" + activeXrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>");
        WriteStreamObject(stream, offsets, 6, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Stale classic trailing page) Tj\nET\n"));

        return stream.ToArray();
    }

    private static byte[] BuildIncrementalClassicXrefPdfWithInheritedTrailerRoot() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        WriteAscii(stream, "%PDF-1.4\n");
        WriteObject(stream, offsets, 1, "<< /Type /Catalog /Pages 2 0 R /PageLayout /TwoColumnLeft >>");
        WriteObject(stream, offsets, 2, "<< /Type /Pages /Count 1 /Kids [3 0 R] >>");
        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 11 0 R >>");
        WriteObject(stream, offsets, 5, "<< /Type /Catalog /Pages 6 0 R /PageLayout /SinglePage >>");
        WriteObject(stream, offsets, 6, "<< /Type /Pages /Count 1 /Kids [7 0 R] >>");
        WriteObject(stream, offsets, 7, "<< /Type /Page /Parent 6 0 R /MediaBox [0 0 200 200] /Contents 11 0 R >>");
        WriteStreamObject(stream, offsets, 11, Array.Empty<byte>());

        int previousXrefOffset = (int)stream.Position;
        var previousEntries = new Dictionary<int, int>(offsets) {
            [0] = 0
        };
        WriteClassicXrefTable(stream, previousEntries, size: 12, rootObjectNumber: 5, previousXrefOffset: null);
        WriteAscii(stream, "startxref\n" + previousXrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        WriteObject(stream, offsets, 8, "<< /Producer (OfficeIMO trailer chain marker) >>");
        int activeXrefOffset = (int)stream.Position;
        var activeEntries = new Dictionary<int, int> {
            [8] = offsets[8]
        };
        WriteClassicXrefTableWithoutRoot(stream, activeEntries, size: 12, previousXrefOffset: previousXrefOffset);
        WriteAscii(stream, "startxref\n" + activeXrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        return stream.ToArray();
    }

    private static byte[] BuildXrefStreamEntries(Dictionary<int, int> offsets, int xrefObjectNumber) {
        using var stream = new MemoryStream();
        WriteXrefEntry(stream, 0, 0, 65535);
        for (int objectNumber = 1; objectNumber <= 8; objectNumber++) {
            if (objectNumber == xrefObjectNumber) {
                WriteXrefEntry(stream, 1, offsets[xrefObjectNumber], 0);
            } else if (offsets.TryGetValue(objectNumber, out int offset)) {
                WriteXrefEntry(stream, 1, offset, 0);
            } else {
                WriteXrefEntry(stream, 0, 0, 65535);
            }
        }

        return stream.ToArray();
    }

    private static byte[] BuildXrefStreamEntries(IReadOnlyDictionary<int, (int Type, int Field1, int Field2)> entries, int size) {
        using var stream = new MemoryStream();
        for (int objectNumber = 0; objectNumber < size; objectNumber++) {
            if (entries.TryGetValue(objectNumber, out var entry)) {
                WriteXrefEntry(stream, entry.Type, entry.Field1, entry.Field2);
            } else {
                WriteXrefEntry(stream, 0, 0, 65535);
            }
        }

        return stream.ToArray();
    }

    private static byte[] BuildXrefStreamEntries(IReadOnlyList<int> objectNumbers, IReadOnlyDictionary<int, (int Type, int Field1, int Field2)> entries) {
        using var stream = new MemoryStream();
        foreach (int objectNumber in objectNumbers) {
            if (!entries.TryGetValue(objectNumber, out var entry)) {
                throw new InvalidOperationException("Missing xref stream entry for object " + objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");
            }

            WriteXrefEntry(stream, entry.Type, entry.Field1, entry.Field2);
        }

        return stream.ToArray();
    }

    private static void WriteClassicXrefTable(Stream stream, IReadOnlyDictionary<int, int> entries, int size, int rootObjectNumber, int? previousXrefOffset) {
        WriteAscii(stream, "xref\n");
        var objectNumbers = entries.Keys.OrderBy(static objectNumber => objectNumber).ToList();
        int index = 0;
        while (index < objectNumbers.Count) {
            int first = objectNumbers[index];
            int end = index + 1;
            while (end < objectNumbers.Count && objectNumbers[end] == objectNumbers[end - 1] + 1) {
                end++;
            }

            WriteAscii(stream, first.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " + (end - index).ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n");
            for (int i = index; i < end; i++) {
                int objectNumber = objectNumbers[i];
                if (objectNumber == 0) {
                    WriteAscii(stream, "0000000000 65535 f \n");
                } else {
                    WriteAscii(stream, entries[objectNumber].ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n");
                }
            }

            index = end;
        }

        WriteAscii(stream, "trailer\n<< /Size " + size.ToString(System.Globalization.CultureInfo.InvariantCulture) + " /Root " + rootObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 R");
        if (previousXrefOffset.HasValue) {
            WriteAscii(stream, " /Prev " + previousXrefOffset.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        WriteAscii(stream, " >>\n");
    }

    private static void WriteClassicXrefTableWithoutRoot(Stream stream, IReadOnlyDictionary<int, int> entries, int size, int? previousXrefOffset) {
        WriteAscii(stream, "xref\n");
        var objectNumbers = entries.Keys.OrderBy(static objectNumber => objectNumber).ToList();
        int index = 0;
        while (index < objectNumbers.Count) {
            int first = objectNumbers[index];
            int end = index + 1;
            while (end < objectNumbers.Count && objectNumbers[end] == objectNumbers[end - 1] + 1) {
                end++;
            }

            WriteAscii(stream, first.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " + (end - index).ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n");
            for (int i = index; i < end; i++) {
                int objectNumber = objectNumbers[i];
                if (objectNumber == 0) {
                    WriteAscii(stream, "0000000000 65535 f \n");
                } else {
                    WriteAscii(stream, entries[objectNumber].ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n");
                }
            }

            index = end;
        }

        WriteAscii(stream, "trailer\n<< /Size " + size.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (previousXrefOffset.HasValue) {
            WriteAscii(stream, " /Prev " + previousXrefOffset.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        WriteAscii(stream, " >>\n");
    }

    private static void WriteXrefEntry(Stream stream, int type, int field1, int field2) {
        stream.WriteByte((byte)type);
        WriteBigEndian(stream, field1, 4);
        WriteBigEndian(stream, field2, 2);
    }

    private static void WriteBigEndian(Stream stream, int value, int byteCount) {
        for (int shift = (byteCount - 1) * 8; shift >= 0; shift -= 8) {
            stream.WriteByte((byte)((value >> shift) & 0xFF));
        }
    }

    private static void WriteObject(Stream stream, Dictionary<int, int> offsets, int objectNumber, string body) {
        offsets[objectNumber] = (int)stream.Position;
        WriteAscii(stream, objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj\n" + body + "\nendobj\n");
    }

    private static void WriteStreamObject(Stream stream, Dictionary<int, int> offsets, int objectNumber, byte[] streamBytes) {
        offsets[objectNumber] = (int)stream.Position;
        WriteAscii(stream, objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Length " +
            streamBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(streamBytes, 0, streamBytes.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");
    }

    private static void WriteRawObject(Stream stream, Dictionary<int, int> offsets, int objectNumber, string objectText) {
        offsets[objectNumber] = (int)stream.Position;
        WriteAscii(stream, objectText);
        if (!objectText.EndsWith("\n", StringComparison.Ordinal)) {
            WriteAscii(stream, "\n");
        }
    }

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static byte[] BuildExternalObjectStreamWithExplicitReplacementPdf() {
        byte[] explicitContent = Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Explicit object wins) Tj\nET\n");
        byte[] packedContent = Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Packed object stream wins) Tj\nET\n");
        var packedObjects = new List<(int ObjectNumber, string Body)> {
            (1, "<< /Type /Catalog /Pages 2 0 R >>"),
            (2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 4 0 R >> >> >>"),
            (3, "<< /Type /Page /Parent 2 0 R /Contents 11 0 R >>"),
            (4, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
        };

        var objects = new[] {
            BuildObjectStreamObject(10, packedObjects),
            BuildStreamObject(11, packedContent),
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, explicitContent)
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildExternalObjectStreamWithCompressedReplacementPdf() {
        byte[] earlierContent = Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Earlier object stream wins) Tj\nET\n");
        byte[] laterContent = Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Later object stream wins) Tj\nET\n");
        var earlierObjects = new List<(int ObjectNumber, string Body)> {
            (1, "<< /Type /Catalog /Pages 2 0 R >>"),
            (2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 4 0 R >> >> >>"),
            (3, "<< /Type /Page /Parent 2 0 R /Contents 11 0 R >>"),
            (4, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
        };
        var laterObjects = new List<(int ObjectNumber, string Body)> {
            (1, "<< /Type /Catalog /Pages 2 0 R >>"),
            (2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 4 0 R >> >> >>"),
            (3, "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>"),
            (4, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
        };

        var objects = new[] {
            BuildObjectStreamObject(10, earlierObjects),
            BuildStreamObject(11, earlierContent),
            BuildObjectStreamObject(12, laterObjects),
            BuildStreamObject(5, laterContent)
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildExternalObjectStreamReplacingEarlierExplicitPdf() {
        byte[] earlierContent = Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Earlier explicit object wins) Tj\nET\n");
        byte[] laterContent = Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Later object stream wins) Tj\nET\n");
        var packedObjects = new List<(int ObjectNumber, string Body)> {
            (1, "<< /Type /Catalog /Pages 2 0 R >>"),
            (2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 4 0 R >> >> >>"),
            (3, "<< /Type /Page /Parent 2 0 R /Contents 11 0 R >>"),
            (4, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
        };

        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, earlierContent),
            BuildObjectStreamObject(10, packedObjects),
            BuildStreamObject(11, laterContent)
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildExternalObjectStreamWithMalformedTrailingPageObjectPdf() {
        byte[] pdf = BuildExternalObjectStreamPdf(includeAcroForm: false);
        byte[] suffix = Encoding.ASCII.GetBytes("\n3 0 obj\nendobj\n");
        var result = new byte[pdf.Length + suffix.Length];
        Buffer.BlockCopy(pdf, 0, result, 0, pdf.Length);
        Buffer.BlockCopy(suffix, 0, result, pdf.Length, suffix.Length);
        return result;
    }

    private static byte[] BuildExternalObjectStreamWithMalformedTrailingPageDictionaryPdf() {
        byte[] pdf = BuildExternalObjectStreamPdf(includeAcroForm: false);
        byte[] suffix = Encoding.ASCII.GetBytes("\n3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 99 0 R /Broken (\nendobj\n");
        var result = new byte[pdf.Length + suffix.Length];
        Buffer.BlockCopy(pdf, 0, result, 0, pdf.Length);
        Buffer.BlockCopy(suffix, 0, result, pdf.Length, suffix.Length);
        return result;
    }

    private static byte[] BuildExternalObjectStreamWithParseFailedTrailingPageDictionaryPdf() {
        byte[] pdf = BuildExternalObjectStreamPdf(includeAcroForm: false);
        byte[] suffix = Encoding.ASCII.GetBytes("\n3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents <ZZ> >>\nendobj\n");
        var result = new byte[pdf.Length + suffix.Length];
        Buffer.BlockCopy(pdf, 0, result, 0, pdf.Length);
        Buffer.BlockCopy(suffix, 0, result, pdf.Length, suffix.Length);
        return result;
    }

    private static byte[] BuildExternalObjectStreamPdf(bool includeAcroForm) {
        byte[] content = Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Object stream page) Tj\nET\n");
        var packedObjects = new List<(int ObjectNumber, string Body)> {
            (1, includeAcroForm ? "<< /Type /Catalog /Pages 2 0 R /AcroForm 6 0 R >>" : "<< /Type /Catalog /Pages 2 0 R >>"),
            (2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 4 0 R >> >> >>"),
            (3, "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>"),
            (4, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
        };

        if (includeAcroForm) {
            packedObjects.Add((6, "<< /Fields [7 0 R] >>"));
            packedObjects.Add((7, "<< /FT /Tx /T (HiddenField) /V (value) >>"));
        }

        var objects = new[] {
            BuildStreamObject(5, content),
            BuildObjectStreamObject(10, packedObjects)
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static string BuildStreamObject(int objectNumber, byte[] streamBytes, string extraDictionary = "") {
        string suffix = string.IsNullOrWhiteSpace(extraDictionary) ? string.Empty : " " + extraDictionary.Trim();
        return objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Length " +
            streamBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            suffix +
            " >>\nstream\n" +
            Encoding.ASCII.GetString(streamBytes) +
            "\nendstream\nendobj";
    }

    private static string BuildObjectStreamObject(int objectNumber, IReadOnlyList<(int ObjectNumber, string Body)> objects) {
        var header = new StringBuilder();
        var body = new StringBuilder();
        for (int i = 0; i < objects.Count; i++) {
            if (i > 0) {
                header.Append(' ');
            }

            header.Append(objects[i].ObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture))
                .Append(' ')
                .Append(body.Length.ToString(System.Globalization.CultureInfo.InvariantCulture));
            body.Append(objects[i].Body).Append('\n');
        }

        header.Append('\n');
        string objectStreamText = header.ToString() + body;
        byte[] encoded = Encoding.ASCII.GetBytes(EncodeAsciiHex(Compress(Encoding.ASCII.GetBytes(objectStreamText))));
        string dictionary = "/Type /ObjStm /N " +
            objects.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /First " +
            Encoding.ASCII.GetByteCount(header.ToString()).ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /Filter [/ASCIIHexDecode /FlateDecode]";
        return BuildStreamObject(objectNumber, encoded, dictionary);
    }

    private static byte[] BuildPdf(IReadOnlyList<string> objects, int rootObjectNumber) {
        var offsets = new Dictionary<int, int>();
        using var stream = new MemoryStream();
        using var writer = new StreamWriter(stream, Encoding.ASCII, 1024, leaveOpen: true);

        writer.WriteLine("%PDF-1.4");
        writer.Flush();
        int maxObjectNumber = 0;
        foreach (string obj in objects) {
            int objectNumber = ReadObjectNumber(obj);
            offsets[objectNumber] = (int)stream.Position;
            maxObjectNumber = Math.Max(maxObjectNumber, objectNumber);
            writer.WriteLine(obj);
            writer.Flush();
        }

        int xrefOffset = (int)stream.Position;
        writer.WriteLine("xref");
        writer.WriteLine("0 " + (maxObjectNumber + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
        writer.WriteLine("0000000000 65535 f ");
        for (int i = 1; i <= maxObjectNumber; i++) {
            if (offsets.TryGetValue(i, out int offset)) {
                writer.WriteLine(offset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n ");
            } else {
                writer.WriteLine("0000000000 65535 f ");
            }
        }

        writer.WriteLine("trailer");
        writer.WriteLine("<< /Size " + (maxObjectNumber + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " /Root " + rootObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 R >>");
        writer.WriteLine("startxref");
        writer.WriteLine(xrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture));
        writer.WriteLine("%%EOF");
        writer.Flush();
        return stream.ToArray();
    }

    private static int ReadObjectNumber(string obj) {
        int space = obj.IndexOf(' ');
        return int.Parse(obj.Substring(0, space), System.Globalization.CultureInfo.InvariantCulture);
    }

    private static byte[] Compress(byte[] input) {
        using var output = new MemoryStream();
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(input, 0, input.Length);
        }

        return output.ToArray();
    }

    private static string EncodeAsciiHex(byte[] bytes) {
        var builder = new StringBuilder(bytes.Length * 2 + 1);
        for (int i = 0; i < bytes.Length; i++) {
            builder.Append(bytes[i].ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
        }

        builder.Append('>');
        return builder.ToString();
    }

    private static string Normalize(string value) {
        return string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }

    private static int CountOccurrences(string value, string text) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(text, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += text.Length;
        }

        return count;
    }
}
