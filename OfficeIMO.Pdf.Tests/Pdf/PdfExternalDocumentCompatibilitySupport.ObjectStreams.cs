using System.IO.Compression;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalDocumentCompatibilityTests {

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

}
