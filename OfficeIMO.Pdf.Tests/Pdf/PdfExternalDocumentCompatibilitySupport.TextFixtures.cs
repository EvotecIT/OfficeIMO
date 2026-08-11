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

    private static byte[] BuildExternalCompactType0ToUnicodePdf() {
        byte[] content = Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n<000100020003000400050006000700080009000A000B000C000D00110008000E000A000F000F0010> Tj\nET\n");
        const string cmap = "\uFEFF/CIDInit /ProcSet findresource begin\n" +
            "12 dict begin\n" +
            "begincmap\n" +
            "/CIDSystemInfo << /Registry (Adobe)/Ordering (UCS)/Supplement 0>> def\n" +
            "/CMapName /Adobe-Identity-UCS def /CMapType 2 def\n" +
            "1 begincodespacerange\n" +
            "<0000><FFFF>\n" +
            "endcodespacerange\n" +
            "17 beginbfrange\n" +
            "<0001><0001><0042>\n" +
            "<0002><0002><0065>\n" +
            "<0003><0003><006E>\n" +
            "<0004><0004><0063>\n" +
            "<0005><0005><0068>\n" +
            "<0006><0006><006D>\n" +
            "<0007><0007><0061>\n" +
            "<0008><0008><0072>\n" +
            "<0009><0009><006B>\n" +
            "<000A><000A><0020>\n" +
            "<000B><000B><0052>\n" +
            "<000C><000C><0065>\n" +
            "<000D><000D><0070>\n" +
            "<000E><000E><0074>\n" +
            "<000F><000F><0030>\n" +
            "<0010><0010><0031>\n" +
            "<0011><0011><006F>\n" +
            "endbfrange\n" +
            "endcmap CMapName currentdict /CMap defineresource pop end end\n";

        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type0 /BaseFont /ExternalSubset /Encoding /Identity-H /DescendantFonts [7 0 R] /ToUnicode 6 0 R >>\nendobj",
            BuildStreamObject(5, content),
            BuildStreamObject(6, Encoding.UTF8.GetBytes(cmap)),
            "7 0 obj\n<< /Type /Font /Subtype /CIDFontType2 /BaseFont /ExternalSubset /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /DW 1000 >>\nendobj"
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildExternalScaledSequentialToUnicodePdf() {
        byte[] content = Encoding.ASCII.GetBytes(
            "BT\n46 0 0 46 72 720 Tm\n/F13 1 Tf\n[(\\)*+,) -1 (%)] TJ\nET\n");
        const string cmap = "/CIDInit /ProcSet findresource begin\n" +
            "12 dict begin\n" +
            "begincmap\n" +
            "4 beginbfrange\n" +
            "<25><25><0065>\n" +
            "<29><29><0054>\n" +
            "<2A><2B><0061>\n" +
            "<2C><2C><006C>\n" +
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
