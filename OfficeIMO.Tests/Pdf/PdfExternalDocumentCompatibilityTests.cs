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
