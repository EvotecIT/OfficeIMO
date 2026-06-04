using System.IO.Compression;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalDocumentCompatibilityTests {

    private static byte[] BuildExternalSinglePagePdf(string content) =>
        BuildExternalSinglePagePdf(content, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");

    private static byte[] BuildExternalSinglePagePdf(string content, string fontObjectBody) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(content.TrimEnd('\n'));
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents [5 0 R] >>\nendobj",
            "4 0 obj\n" + fontObjectBody + "\nendobj",
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

    private static byte[] BuildExternalThreeByteToUnicodePdf() {
        byte[] content = Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n<010203> Tj\nET\n");
        const string cmap = "/CIDInit /ProcSet findresource begin\n" +
            "12 dict begin\n" +
            "begincmap\n" +
            "1 beginbfchar\n" +
            "<010203> <005A>\n" +
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

}
