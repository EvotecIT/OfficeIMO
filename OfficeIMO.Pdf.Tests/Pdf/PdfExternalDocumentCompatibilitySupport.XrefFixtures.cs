using System.IO.Compression;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalDocumentCompatibilityTests {

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

    private static byte[] BuildXrefStreamPdfWithTrailingStaleObjectStreamPage() {
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

        WriteStreamObject(stream, offsets, 6, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Stale trailing object stream page) Tj\nET\n"));
        var stalePackedObjects = new List<(int ObjectNumber, string Body)> {
            (3, "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>")
        };
        WriteRawObject(stream, offsets, 20, BuildObjectStreamObject(20, stalePackedObjects));

        return stream.ToArray();
    }

    private static byte[] BuildClassicXrefPdfWithTrailingStaleObjectStreamPage() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        WriteAscii(stream, "%PDF-1.4\n");
        WriteObject(stream, offsets, 1, "<< /Type /Catalog /Pages 2 0 R >>");
        WriteObject(stream, offsets, 2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 7 0 R >> >> >>");
        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>");
        WriteStreamObject(stream, offsets, 4, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Active classic xref page) Tj\nET\n"));
        WriteObject(stream, offsets, 7, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");

        int classicXrefOffset = (int)stream.Position;
        var classicEntries = new Dictionary<int, int>(offsets) {
            [0] = 0
        };
        WriteClassicXrefTable(stream, classicEntries, size: 21, rootObjectNumber: 1, previousXrefOffset: null);
        WriteAscii(stream, "startxref\n" + classicXrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        WriteStreamObject(stream, offsets, 6, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Stale classic object stream page) Tj\nET\n"));
        var stalePackedObjects = new List<(int ObjectNumber, string Body)> {
            (3, "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>")
        };
        WriteRawObject(stream, offsets, 20, BuildObjectStreamObject(20, stalePackedObjects));

        return stream.ToArray();
    }

    internal static byte[] BuildIncrementalClassicXrefPdfWithWrongGenerationReplacementPage() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        WriteAscii(stream, "%PDF-1.4\n");
        WriteObject(stream, offsets, 1, "<< /Type /Catalog /Pages 2 0 R >>");
        WriteObject(stream, offsets, 2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 7 0 R >> >> >>");
        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>");
        WriteStreamObject(stream, offsets, 4, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Previous classic generation page) Tj\nET\n"));
        WriteObject(stream, offsets, 7, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");

        int previousXrefOffset = (int)stream.Position;
        var previousEntries = new Dictionary<int, int>(offsets) {
            [0] = 0
        };
        WriteClassicXrefTable(stream, previousEntries, size: 9, rootObjectNumber: 1, previousXrefOffset: null);
        WriteAscii(stream, "startxref\n" + previousXrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        WriteObject(stream, offsets, 6, "<< /Producer (wrong generation marker) >>");
        WriteObjectGeneration(stream, offsets, 3, 1, "<< /Type /Page /Parent 2 0 R /Contents 8 0 R >>");
        WriteStreamObject(stream, offsets, 8, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Wrong classic generation page) Tj\nET\n"));

        int activeXrefOffset = (int)stream.Position;
        var activeEntries = new Dictionary<int, int> {
            [3] = offsets[3],
            [6] = offsets[6],
            [8] = offsets[8]
        };
        WriteClassicXrefTable(stream, activeEntries, size: 9, rootObjectNumber: 1, previousXrefOffset: previousXrefOffset);
        WriteAscii(stream, "startxref\n" + activeXrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        return stream.ToArray();
    }

    private static byte[] BuildIncrementalXrefStreamPdfWithWrongGenerationReplacementPage() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        WriteAscii(stream, "%PDF-1.5\n");
        WriteObject(stream, offsets, 1, "<< /Type /Catalog /Pages 2 0 R >>");
        WriteObject(stream, offsets, 2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 7 0 R >> >> >>");
        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>");
        WriteStreamObject(stream, offsets, 4, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Previous xref generation page) Tj\nET\n"));
        WriteObject(stream, offsets, 7, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");

        int previousXrefObjectNumber = 10;
        offsets[previousXrefObjectNumber] = (int)stream.Position;
        var previousEntries = new Dictionary<int, (int Type, int Field1, int Field2)>();
        foreach (var offset in offsets) {
            previousEntries[offset.Key] = (1, offset.Value, 0);
        }

        byte[] previousXrefEntries = BuildXrefStreamEntries(previousEntries, size: 12);
        WriteAscii(stream, previousXrefObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 12 /Root 1 0 R /W [1 4 2] /Index [0 12] /Length " +
            previousXrefEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(previousXrefEntries, 0, previousXrefEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");
        WriteAscii(stream, "startxref\n" + offsets[previousXrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        WriteObject(stream, offsets, 6, "<< /Producer (wrong generation marker) >>");
        WriteObjectGeneration(stream, offsets, 3, 1, "<< /Type /Page /Parent 2 0 R /Contents 8 0 R >>");
        WriteStreamObject(stream, offsets, 8, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Wrong xref generation page) Tj\nET\n"));

        int activeXrefObjectNumber = 11;
        offsets[activeXrefObjectNumber] = (int)stream.Position;
        var activeEntries = new Dictionary<int, (int Type, int Field1, int Field2)> {
            [3] = (1, offsets[3], 0),
            [6] = (1, offsets[6], 0),
            [8] = (1, offsets[8], 0),
            [11] = (1, offsets[activeXrefObjectNumber], 0)
        };
        byte[] xrefEntries = BuildXrefStreamEntries(new[] { 3, 6, 8, 11 }, activeEntries);
        WriteAscii(stream, activeXrefObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 12 /Root 1 0 R /Prev " +
            offsets[previousXrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /W [1 4 2] /Index [3 1 6 1 8 1 11 1] /Length " +
            xrefEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(xrefEntries, 0, xrefEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");
        WriteAscii(stream, "startxref\n" + offsets[activeXrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        return stream.ToArray();
    }

    private static byte[] BuildClassicXrefPdfWithWrongGenerationContentReference() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        WriteAscii(stream, "%PDF-1.4\n");
        WriteObject(stream, offsets, 1, "<< /Type /Catalog /Pages 2 0 R >>");
        WriteObject(stream, offsets, 2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 7 0 R >> >> >>");
        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>");
        WriteStreamObjectGeneration(stream, offsets, 4, 1, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Wrong generation referenced content) Tj\nET\n"));
        WriteObject(stream, offsets, 7, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");

        int xrefOffset = (int)stream.Position;
        var entries = new Dictionary<int, int>(offsets) {
            [0] = 0
        };
        WriteClassicXrefTable(stream, entries, size: 8, rootObjectNumber: 1, previousXrefOffset: null);
        WriteAscii(stream, "startxref\n" + xrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

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

    private static byte[] BuildIncrementalXrefStreamPdfWithClassicPrevAndTrailingStaleDuplicatePage() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        WriteAscii(stream, "%PDF-1.5\n");
        WriteObject(stream, offsets, 1, "<< /Type /Catalog /Pages 2 0 R >>");
        WriteObject(stream, offsets, 2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 7 0 R >> >> >>");
        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>");
        WriteStreamObject(stream, offsets, 4, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Inherited mixed xref page) Tj\nET\n"));
        WriteObject(stream, offsets, 7, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");

        int previousXrefOffset = (int)stream.Position;
        var previousEntries = new Dictionary<int, int>(offsets) {
            [0] = 0
        };
        WriteClassicXrefTable(stream, previousEntries, size: 9, rootObjectNumber: 1, previousXrefOffset: null);
        WriteAscii(stream, "startxref\n" + previousXrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        int activeXrefObjectNumber = 8;
        offsets[activeXrefObjectNumber] = (int)stream.Position;
        var activeEntries = new Dictionary<int, (int Type, int Field1, int Field2)> {
            [8] = (1, offsets[activeXrefObjectNumber], 0)
        };
        byte[] xrefEntries = BuildXrefStreamEntries(new[] { 8 }, activeEntries);
        WriteAscii(stream, activeXrefObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 9 /Root 1 0 R /Prev " +
            previousXrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /W [1 4 2] /Index [8 1] /Length " +
            xrefEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(xrefEntries, 0, xrefEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");
        WriteAscii(stream, "startxref\n" + offsets[activeXrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>");
        WriteStreamObject(stream, offsets, 6, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Stale mixed xref trailing page) Tj\nET\n"));

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

    internal static byte[] BuildHybridClassicXrefPdfWithXRefStmAndTrailingStaleDuplicatePage() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        WriteAscii(stream, "%PDF-1.5\n");
        WriteObject(stream, offsets, 1, "<< /Type /Catalog /Pages 2 0 R >>");
        WriteObject(stream, offsets, 2, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 7 0 R >> >> >>");
        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>");
        WriteStreamObject(stream, offsets, 4, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Hybrid xref stream page) Tj\nET\n"));
        WriteObject(stream, offsets, 7, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");

        int xrefStreamObjectNumber = 8;
        offsets[xrefStreamObjectNumber] = (int)stream.Position;
        var xrefStreamEntries = new Dictionary<int, (int Type, int Field1, int Field2)> {
            [3] = (1, offsets[3], 0)
        };
        byte[] xrefEntries = BuildXrefStreamEntries(new[] { 3 }, xrefStreamEntries);
        WriteAscii(stream, xrefStreamObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 9 /Root 1 0 R /W [1 4 2] /Index [3 1] /Length " +
            xrefEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(xrefEntries, 0, xrefEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");

        int classicXrefOffset = (int)stream.Position;
        var classicEntries = new Dictionary<int, int> {
            [0] = 0,
            [1] = offsets[1],
            [2] = offsets[2],
            [4] = offsets[4],
            [7] = offsets[7],
            [8] = offsets[8]
        };
        WriteClassicXrefTableWithXRefStm(stream, classicEntries, size: 9, rootObjectNumber: 1, xrefStreamOffset: offsets[xrefStreamObjectNumber]);
        WriteAscii(stream, "startxref\n" + classicXrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>");
        WriteStreamObject(stream, offsets, 6, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Stale hybrid trailing page) Tj\nET\n"));

        return stream.ToArray();
    }

    private static byte[] BuildHybridClassicXrefPdfWithXRefStmTrailerRootAndStaleXrefStreamRoot() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        WriteAscii(stream, "%PDF-1.5\n");
        WriteObject(stream, offsets, 1, "<< /Type /Catalog /Pages 2 0 R /PageLayout /TwoColumnLeft >>");
        WriteObject(stream, offsets, 2, "<< /Type /Pages /Count 1 /Kids [3 0 R] >>");
        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 11 0 R >>");
        WriteObject(stream, offsets, 5, "<< /Type /Catalog /Pages 6 0 R /PageLayout /SinglePage >>");
        WriteObject(stream, offsets, 6, "<< /Type /Pages /Count 1 /Kids [7 0 R] >>");
        WriteObject(stream, offsets, 7, "<< /Type /Page /Parent 6 0 R /MediaBox [0 0 200 200] /Contents 11 0 R >>");
        WriteStreamObject(stream, offsets, 11, Array.Empty<byte>());

        int xrefStreamObjectNumber = 12;
        offsets[xrefStreamObjectNumber] = (int)stream.Position;
        var xrefStreamEntries = new Dictionary<int, (int Type, int Field1, int Field2)> {
            [12] = (1, offsets[xrefStreamObjectNumber], 0)
        };
        byte[] xrefEntries = BuildXrefStreamEntries(new[] { 12 }, xrefStreamEntries);
        WriteAscii(stream, xrefStreamObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 13 /Root 5 0 R /W [1 4 2] /Index [12 1] /Length " +
            xrefEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(xrefEntries, 0, xrefEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");

        int classicXrefOffset = (int)stream.Position;
        var classicEntries = new Dictionary<int, int>(offsets) {
            [0] = 0
        };
        WriteClassicXrefTableWithXRefStmWithoutRoot(stream, classicEntries, size: 13, xrefStreamOffset: offsets[xrefStreamObjectNumber]);
        WriteAscii(stream, "startxref\n" + classicXrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        int staleXrefObjectNumber = 50;
        offsets[staleXrefObjectNumber] = (int)stream.Position;
        var staleEntries = new Dictionary<int, (int Type, int Field1, int Field2)> {
            [50] = (1, offsets[staleXrefObjectNumber], 0)
        };
        byte[] staleXrefEntries = BuildXrefStreamEntries(new[] { 50 }, staleEntries);
        WriteAscii(stream, staleXrefObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 51 /Root 1 0 R /W [1 4 2] /Index [50 1] /Length " +
            staleXrefEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(staleXrefEntries, 0, staleXrefEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");

        return stream.ToArray();
    }

    private static byte[] BuildIncrementalXrefStreamPdfWithInheritedTrailerRootAndStaleHighObjectXref() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        WriteAscii(stream, "%PDF-1.5\n");
        WriteObject(stream, offsets, 1, "<< /Type /Catalog /Pages 2 0 R /PageLayout /TwoColumnLeft >>");
        WriteObject(stream, offsets, 2, "<< /Type /Pages /Count 1 /Kids [3 0 R] >>");
        WriteObject(stream, offsets, 3, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 11 0 R >>");
        WriteObject(stream, offsets, 5, "<< /Type /Catalog /Pages 6 0 R /PageLayout /SinglePage >>");
        WriteObject(stream, offsets, 6, "<< /Type /Pages /Count 1 /Kids [7 0 R] >>");
        WriteObject(stream, offsets, 7, "<< /Type /Page /Parent 6 0 R /MediaBox [0 0 200 200] /Contents 11 0 R >>");
        WriteStreamObject(stream, offsets, 11, Array.Empty<byte>());

        int previousXrefObjectNumber = 12;
        offsets[previousXrefObjectNumber] = (int)stream.Position;
        var previousEntries = new Dictionary<int, (int Type, int Field1, int Field2)> {
            [1] = (1, offsets[1], 0),
            [2] = (1, offsets[2], 0),
            [3] = (1, offsets[3], 0),
            [5] = (1, offsets[5], 0),
            [6] = (1, offsets[6], 0),
            [7] = (1, offsets[7], 0),
            [11] = (1, offsets[11], 0),
            [12] = (1, offsets[previousXrefObjectNumber], 0)
        };
        byte[] previousXrefEntries = BuildXrefStreamEntries(previousEntries, size: 13);
        WriteAscii(stream, previousXrefObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 13 /Root 5 0 R /W [1 4 2] /Index [0 13] /Length " +
            previousXrefEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(previousXrefEntries, 0, previousXrefEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");
        WriteAscii(stream, "startxref\n" + offsets[previousXrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        int activeXrefObjectNumber = 13;
        offsets[activeXrefObjectNumber] = (int)stream.Position;
        var activeEntries = new Dictionary<int, (int Type, int Field1, int Field2)> {
            [13] = (1, offsets[activeXrefObjectNumber], 0)
        };
        byte[] activeXrefEntries = BuildXrefStreamEntries(new[] { 13 }, activeEntries);
        WriteAscii(stream, activeXrefObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 14 /Prev " +
            offsets[previousXrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /W [1 4 2] /Index [13 1] /Length " +
            activeXrefEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(activeXrefEntries, 0, activeXrefEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");
        WriteAscii(stream, "startxref\n" + offsets[activeXrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        int staleXrefObjectNumber = 50;
        offsets[staleXrefObjectNumber] = (int)stream.Position;
        var staleEntries = new Dictionary<int, (int Type, int Field1, int Field2)> {
            [50] = (1, offsets[staleXrefObjectNumber], 0)
        };
        byte[] staleXrefEntries = BuildXrefStreamEntries(new[] { 50 }, staleEntries);
        WriteAscii(stream, staleXrefObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 51 /Root 1 0 R /W [1 4 2] /Index [50 1] /Length " +
            staleXrefEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(staleXrefEntries, 0, staleXrefEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");

        return stream.ToArray();
    }

    private static byte[] BuildIncrementalXrefStreamPdfWithClassicTrailerRoot() {
        using var stream = new MemoryStream();
        var offsets = new Dictionary<int, int>();

        WriteAscii(stream, "%PDF-1.5\n");
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
        WriteClassicXrefTable(stream, previousEntries, size: 13, rootObjectNumber: 5, previousXrefOffset: null);
        WriteAscii(stream, "startxref\n" + previousXrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        int activeXrefObjectNumber = 12;
        offsets[activeXrefObjectNumber] = (int)stream.Position;
        var activeEntries = new Dictionary<int, (int Type, int Field1, int Field2)> {
            [12] = (1, offsets[activeXrefObjectNumber], 0)
        };
        byte[] activeXrefEntries = BuildXrefStreamEntries(new[] { 12 }, activeEntries);
        WriteAscii(stream, activeXrefObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Type /XRef /Size 13 /Prev " +
            previousXrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /W [1 4 2] /Index [12 1] /Length " +
            activeXrefEntries.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(activeXrefEntries, 0, activeXrefEntries.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");
        WriteAscii(stream, "startxref\n" + offsets[activeXrefObjectNumber].ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n%%EOF\n");

        return stream.ToArray();
    }

}
