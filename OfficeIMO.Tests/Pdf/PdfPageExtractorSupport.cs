using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageExtractorTests {
    private static byte[] BuildSinglePagePdfWithGenerationOneContent(
        bool includeAnnotation = false,
        int contentObjectGeneration = 1,
        int contentReferenceGeneration = 1,
        int annotationObjectGeneration = 1,
        int annotationReferenceGeneration = 1) {
        var stream = new MemoryStream();
        var offsets = new Dictionary<int, (int Offset, int Generation)>();
        string annotationEntry = includeAnnotation ? " /Annots [6 " + annotationReferenceGeneration.ToString(System.Globalization.CultureInfo.InvariantCulture) + " R]" : string.Empty;

        WriteAscii(stream, "%PDF-1.4\n");
        WriteObject(stream, offsets, 1, 0, "<< /Type /Catalog /Pages 2 0 R >>");
        WriteObject(stream, offsets, 2, 0, "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 5 0 R >> >> >>");
        WriteObject(stream, offsets, 3, 0, "<< /Type /Page /Parent 2 0 R /Contents 4 " + contentReferenceGeneration.ToString(System.Globalization.CultureInfo.InvariantCulture) + " R" + annotationEntry + " >>");
        WriteStreamObject(stream, offsets, 4, contentObjectGeneration, Encoding.ASCII.GetBytes("BT\n/F13 12 Tf\n72 720 Td\n(Generation one content) Tj\nET\n"));
        WriteObject(stream, offsets, 5, 0, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");

        if (includeAnnotation) {
            WriteObject(stream, offsets, 6, annotationObjectGeneration, "<< /Type /Annot /Subtype /Link /Rect [72 700 180 722] /Border [0 0 0] /A << /S /URI /URI (https://evotec.xyz) >> /Contents (Generation link) >>");
        }

        int size = includeAnnotation ? 7 : 6;
        int xrefOffset = (int)stream.Position;
        WriteAscii(stream, "xref\n0 " + size.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n");
        WriteAscii(stream, "0000000000 65535 f \n");

        for (int objectNumber = 1; objectNumber < size; objectNumber++) {
            var entry = offsets[objectNumber];
            WriteAscii(stream,
                entry.Offset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture) +
                " " +
                entry.Generation.ToString("D5", System.Globalization.CultureInfo.InvariantCulture) +
                " n \n");
        }

        WriteAscii(stream,
            "trailer\n<< /Size " + size.ToString(System.Globalization.CultureInfo.InvariantCulture) + " /Root 1 0 R >>\n" +
            "startxref\n" +
            xrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            "\n%%EOF\n");

        return stream.ToArray();
    }

    private static void WriteObject(Stream stream, Dictionary<int, (int Offset, int Generation)> offsets, int objectNumber, int generation, string body) {
        offsets[objectNumber] = ((int)stream.Position, generation);
        WriteAscii(stream, objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " + generation.ToString(System.Globalization.CultureInfo.InvariantCulture) + " obj\n");
        WriteAscii(stream, body);
        WriteAscii(stream, "\nendobj\n");
    }

    private static void WriteStreamObject(Stream stream, Dictionary<int, (int Offset, int Generation)> offsets, int objectNumber, int generation, byte[] streamBytes) {
        offsets[objectNumber] = ((int)stream.Position, generation);
        WriteAscii(stream, objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " + generation.ToString(System.Globalization.CultureInfo.InvariantCulture) + " obj\n");
        WriteAscii(stream, "<< /Length " + streamBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        stream.Write(streamBytes, 0, streamBytes.Length);
        WriteAscii(stream, "endstream\nendobj\n");
    }

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static byte[] BuildThreePagePdf() {
        var doc = PdfDocument.Create()
            .Meta(
                title: "Extraction sample",
                author: "OfficeIMO",
                subject: "Manipulation",
                keywords: "pdf,extract,split");

        doc.Compose(compose => {
            compose.Page(page => {
                page.Size(PageSizes.A4);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text(PageMarker(1)))));
            });

            compose.Page(page => {
                page.Size(new PageSize(612, 792));
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text(PageMarker(2)))));
            });

            compose.Page(page => {
                page.Size(new PageSize(300, 500));
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text(PageMarker(3)))));
            });
        });

        return doc.ToBytes();
    }

    private static string PageMarker(int pageNumber) {
        return pageNumber switch {
            1 => "First page marker",
            2 => "Second page marker",
            3 => "Third page marker",
            _ => "Page " + pageNumber
        };
    }

    private static string NormalizeExtractedText(string text) {
        return text.Replace(" ", string.Empty);
    }

    private static void AssertContainsInOrder(string text, params string[] expected) {
        int previous = -1;
        foreach (string item in expected) {
            int index = text.IndexOf(item, previous + 1, StringComparison.Ordinal);
            Assert.True(index >= 0, "Expected text '" + item + "' was not found after index " + previous + " in '" + text + "'.");
            previous = index;
        }
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += value.Length;
        }

        return count;
    }

    private static bool TableContainsRow(PdfLogicalTable table, params string[] expectedCells) {
        return table.Rows.Any(row => RowContains(row, expectedCells));
    }

    private static bool RowContains(IReadOnlyList<string> row, params string[] expectedCells) {
        return expectedCells.All(expected => row.Any(cell => NormalizeExtractedText(cell).Contains(expected, StringComparison.Ordinal)));
    }

    private static byte[] CreateMinimalRgbPng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 12,
            73, 68, 65, 84,
            0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }

    private sealed class WriteOnlyStream : MemoryStream {
        public override bool CanRead => false;
    }

    private sealed class ReadOnlyStream : MemoryStream {
        public override bool CanWrite => false;
    }
}
