using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfStamperTests {
    private static byte[] BuildTwoPagePdf() {
        var doc = PdfDocument.Create()
            .Meta(title: "Stamp sample", author: "OfficeIMO");

        doc.Compose(compose => {
            compose.Page(page => {
                page.Size(PageSizes.A4);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("First page body"))));
            });

            compose.Page(page => {
                page.Size(PageSizes.A4);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Second page body"))));
            });
        });

        return doc.ToBytes();
    }

    private static byte[] BuildIndirectContentsArrayPdf() {
        string first = "BT /F1 12 Tf 20 80 Td (First) Tj ET";
        string second = "BT /F1 12 Tf 20 60 Td (Second) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 8 0 R /Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> >>",
            "endobj",
            "4 0 obj",
            $"<< /Length {first.Length} >>",
            "stream",
            first,
            "endstream",
            "endobj",
            "5 0 obj",
            $"<< /Length {second.Length} >>",
            "stream",
            second,
            "endstream",
            "endobj",
            "8 0 obj",
            "[4 0 R 5 0 R]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static string Normalize(string text) {
        return text.Replace(" ", string.Empty);
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

    private static MemoryStream CreatePrefixedStream(byte[] pdf) {
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        var stream = new MemoryStream();
        stream.Write(prefix, 0, prefix.Length);
        stream.Write(pdf, 0, pdf.Length);
        stream.Position = prefix.Length;
        return stream;
    }

    private static MemoryStream CreateOutputStream(out int prefixLength) {
        byte[] prefix = Encoding.ASCII.GetBytes("output-prefix");
        var stream = new MemoryStream();
        stream.Write(prefix, 0, prefix.Length);
        prefixLength = prefix.Length;
        return stream;
    }

    private static byte[] GetOutputPayload(MemoryStream output, int prefixLength) {
        byte[] bytes = output.ToArray();
        Assert.True(bytes.Length > prefixLength);
        Assert.Equal("output-prefix", Encoding.ASCII.GetString(bytes, 0, prefixLength));

        var payload = new byte[bytes.Length - prefixLength];
        Array.Copy(bytes, prefixLength, payload, 0, payload.Length);
        return payload;
    }

    private static string FindContentStreamContaining(byte[] pdf, string marker) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        foreach (var item in objects.Values) {
            if (item.Value is PdfStream stream) {
                string content = DecodeStream(stream.Data);
                if (content.Contains(marker)) {
                    return content;
                }
            }
        }

        throw new InvalidOperationException("Content stream marker was not found: " + marker);
    }

    private static IReadOnlyList<string> GetPageContentStreams(byte[] pdf, int pageNumber) {
        var document = PdfReadDocument.Open(pdf);
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        int pageObjectNumber = document.Pages[pageNumber - 1].ObjectNumber;
        if (!objects.TryGetValue(pageObjectNumber, out var pageObject) || pageObject.Value is not PdfDictionary pageDictionary) {
            throw new InvalidOperationException("Page object was not found.");
        }

        if (!pageDictionary.Items.TryGetValue("Contents", out var contents)) {
            throw new InvalidOperationException("Page contents were not found.");
        }

        var streams = new List<string>();
        AppendContentStreams(objects, contents, streams);
        return streams;
    }

    private static void AppendContentStreams(Dictionary<int, PdfIndirectObject> objects, PdfObject contents, List<string> streams) {
        if (contents is PdfReference reference) {
            if (objects.TryGetValue(reference.ObjectNumber, out var indirect) && indirect.Value is PdfStream stream) {
                streams.Add(DecodeStream(stream.Data));
            }

            return;
        }

        if (contents is PdfArray array) {
            foreach (var item in array.Items) {
                AppendContentStreams(objects, item, streams);
            }
        }
    }

    private static string DecodeStream(byte[] data) {
        return Encoding.GetEncoding("ISO-8859-1").GetString(data);
    }

    private sealed class WriteOnlyStream : MemoryStream {
        public override bool CanRead => false;
    }

    private sealed class ReadOnlyStream : MemoryStream {
        public override bool CanWrite => false;
    }

    private static byte[] CreateMinimalRgbPng() => PdfPngTestImages.CreateRgbPng(255, 0, 0);

    private static byte[] CreateMinimalRgbaPng() => PdfPngTestImages.CreateRgbaPng(255, 0, 0, 128);

    private static byte[] CreateMinimalIndexedColorPng() {
        using var ms = new MemoryStream();
        byte[] signature = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
        ms.Write(signature, 0, signature.Length);
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 2,
            0, 0, 0, 1,
            8, 3, 0, 0, 0
        });
        WritePngChunk(ms, "PLTE", new byte[] {
            0xE6, 0x39, 0x46,
            0x2B, 0x7D, 0xD8
        });
        WritePngChunk(ms, "tRNS", new byte[] { 255, 64 });
        WritePngChunk(ms, "IDAT", new byte[] {
            0x78, 0x01,
            0x01,
            0x03, 0x00,
            0xFC, 0xFF,
            0x00, 0x00, 0x01,
            0x00, 0x04, 0x00, 0x02
        });
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    private static byte[] CreateMinimalRgbTransparencyPng() {
        using var ms = new MemoryStream();
        byte[] signature = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
        ms.Write(signature, 0, signature.Length);
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0
        });
        WritePngChunk(ms, "tRNS", new byte[] {
            0, 255,
            0, 0,
            0, 0
        });
        WritePngChunk(ms, "IDAT", new byte[] {
            0x78, 0x01,
            0x01,
            0x04, 0x00,
            0xFB, 0xFF,
            0x00, 0xFF, 0x00, 0x00,
            0x03, 0x01, 0x01, 0x00
        });
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    private static byte[] CreateMinimalPackedGrayscalePng() {
        using var ms = new MemoryStream();
        byte[] signature = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
        ms.Write(signature, 0, signature.Length);
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 2,
            0, 0, 0, 1,
            4, 0, 0, 0, 0
        });
        WritePngChunk(ms, "tRNS", new byte[] { 0, 1 });
        WritePngChunk(ms, "IDAT", new byte[] {
            0x78, 0x01,
            0x01,
            0x02, 0x00,
            0xFD, 0xFF,
            0x00, 0x01,
            0x00, 0x03, 0x00, 0x02
        });
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    private static void WritePngChunk(Stream stream, string type, byte[] data) {
        stream.WriteByte((byte)((data.Length >> 24) & 0xFF));
        stream.WriteByte((byte)((data.Length >> 16) & 0xFF));
        stream.WriteByte((byte)((data.Length >> 8) & 0xFF));
        stream.WriteByte((byte)(data.Length & 0xFF));
        byte[] typeBytes = Encoding.ASCII.GetBytes(type);
        stream.Write(typeBytes, 0, typeBytes.Length);
        stream.Write(data, 0, data.Length);
        uint crc = ComputeCrc32(typeBytes, data);
        stream.WriteByte((byte)((crc >> 24) & 0xFF));
        stream.WriteByte((byte)((crc >> 16) & 0xFF));
        stream.WriteByte((byte)((crc >> 8) & 0xFF));
        stream.WriteByte((byte)(crc & 0xFF));
    }

    private static uint ComputeCrc32(byte[] typeBytes, byte[] data) {
        uint crc = 0xFFFFFFFF;
        for (int i = 0; i < typeBytes.Length; i++) {
            crc = UpdateCrc32(crc, typeBytes[i]);
        }

        for (int i = 0; i < data.Length; i++) {
            crc = UpdateCrc32(crc, data[i]);
        }

        return crc ^ 0xFFFFFFFF;
    }

    private static uint UpdateCrc32(uint crc, byte value) {
        crc ^= value;
        for (int bit = 0; bit < 8; bit++) {
            crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320 : crc >> 1;
        }

        return crc;
    }

}
