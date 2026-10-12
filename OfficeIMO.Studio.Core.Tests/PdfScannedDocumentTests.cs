using System.IO.Compression;
using System.Text;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfScannedDocumentTests {
    [Theory]
    [InlineData(6, 2550, 3300)] // Letter at 300 dpi in colour: each page image decodes to 24 MiB.
    [InlineData(2, 3508, 4961)] // A3 at 300 dpi in colour: 50 MiB and 17.4 million pixels.
    public async Task RendersEveryPageOfA300DpiColourScan(int pages, int width, int height) {
        string path = Path.Combine(Path.GetTempPath(), $"officeimo-scan-{Guid.NewGuid():N}.pdf");
        await File.WriteAllBytesAsync(path, CreateScan(pages, width, height));
        try {
            using var workspace = await PdfWorkspace.OpenAsync(path, default);
            PdfDocumentSession session = PdfDocumentSession.FromWorkspace(workspace);
            Assert.Equal(pages, session.Pages.Count);
            for (int page = 1; page <= pages; page++) {
                PdfRenderedPage rendered = await session.RenderPageAsync(page, 2D, default);
                Assert.Equal(1224, rendered.PixelWidth);
                Assert.Equal(1584, rendered.PixelHeight);
            }
        } finally { File.Delete(path); }
    }

    private static byte[] CreateScan(int pages, int width, int height) {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            $"<< /Type /Pages /Count {pages} /Kids [{string.Join(' ', Enumerable.Range(0, pages).Select(page => $"{3 + page * 3} 0 R"))}] >>"
        };
        var streams = new Dictionary<int, byte[]>();
        for (int page = 0; page < pages; page++) {
            int pageObject = 3 + page * 3;
            objects.Add($"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /XObject << /Im0 {pageObject + 1} 0 R >> >> /Contents {pageObject + 2} 0 R >>");
            // Each page has its own image, as a scan does; a varying band keeps the images distinct.
            byte[] pixels = new byte[width * height * 3];
            Array.Fill(pixels, (byte)0xF0);
            pixels.AsSpan(page * width * 3 * 10, width * 3 * 10).Fill(0x20);
            byte[] image = Compress(pixels);
            streams[objects.Count + 1] = image;
            objects.Add($"<< /Type /XObject /Subtype /Image /Width {width} /Height {height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length {image.Length} >>");
            byte[] content = Encoding.ASCII.GetBytes("q 612 0 0 792 0 0 cm /Im0 Do Q");
            streams[objects.Count + 1] = content;
            objects.Add($"<< /Length {content.Length} >>");
        }

        using var output = new MemoryStream();
        void Write(string text) => output.Write(Encoding.ASCII.GetBytes(text));
        Write("%PDF-1.7\n");
        var offsets = new List<long>();
        for (int index = 0; index < objects.Count; index++) {
            int number = index + 1;
            offsets.Add(output.Position);
            Write($"{number} 0 obj\n{objects[index]}\n");
            if (streams.TryGetValue(number, out byte[]? stream)) {
                Write("stream\n");
                output.Write(stream);
                Write("\nendstream\n");
            }
            Write("endobj\n");
        }
        long xref = output.Position;
        Write($"xref\n0 {objects.Count + 1}\n0000000000 65535 f \n");
        foreach (long offset in offsets) Write($"{offset:D10} 00000 n \n");
        Write($"trailer << /Root 1 0 R /Size {objects.Count + 1} >>\nstartxref\n{xref}\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] Compress(byte[] data) {
        using var output = new MemoryStream();
        using (var zlib = new ZLibStream(output, CompressionLevel.Fastest)) zlib.Write(data);
        return output.ToArray();
    }
}
