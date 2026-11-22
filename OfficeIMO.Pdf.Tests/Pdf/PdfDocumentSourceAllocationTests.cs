#if NET8_0_OR_GREATER && PDF_PERFORMANCE_EVIDENCE
using OfficeIMO.Pdf;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Tests.Pdf;

// Opt-in evidence: build with -p:PdfPerformanceEvidence=true and run --filter Category=Performance.
public sealed class PdfDocumentSourceAllocationTests {
    private readonly ITestOutputHelper _output;

    public PdfDocumentSourceAllocationTests(ITestOutputHelper output) => _output = output;

    [Fact]
    [Trait("Category", "ResourcePerformanceEvidence")]
    [Trait("Category", "Performance")]
    public void LoadAndFirstPageRenderAllocationsByPdfSource() {
        byte[] text = PdfDocumentSourceLifetimeTests.CreateTextPdf(pageCount: 2_000);
        byte[] images = CreateImagePdf(pageCount: 20);

        _output.WriteLine("| File | Source | File bytes | Load | Render page 1 |");
        _output.WriteLine("| --- | --- | ---: | ---: | ---: |");
        foreach ((string name, byte[] pdf) in new[] { ("2,000 text pages", text), ("20 image pages", images) }) {
            long streamLoad = Measure(name, "stream", pdf, bytes => PdfDocument.Load(new MemoryStream(bytes, writable: false)));
            long arrayLoad = Measure(name, "byte[]", pdf, bytes => PdfDocument.Load(bytes));

            // The array is read in place, so loading never pays for another copy of the file.
            Assert.True(arrayLoad < streamLoad - pdf.Length / 2, name + ": byte[] " + arrayLoad + " vs stream " + streamLoad);
        }
    }

    private long Measure(string name, string source, byte[] pdf, Func<byte[], PdfDocument> load) {
        long loadBytes = 0, renderBytes = 0;
        for (int attempt = 0; attempt < 3; attempt++) {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long before = GC.GetAllocatedBytesForCurrentThread();
            using PdfDocument document = load(pdf);
            long loaded = GC.GetAllocatedBytesForCurrentThread();
            Assert.Single(document.ToImages(new PdfImageExportOptions { PageSelector = OfficePageSelector.Parse("1") }));
            long rendered = GC.GetAllocatedBytesForCurrentThread();
            loadBytes = loaded - before;
            renderBytes = rendered - loaded;
        }

        _output.WriteLine("| " + name + " | " + source + " | " + pdf.Length.ToString("N0") + " | " +
            loadBytes.ToString("N0") + " (" + (loadBytes / (double)pdf.Length).ToString("F2") + "x) | " + renderBytes.ToString("N0") + " |");
        return loadBytes;
    }

    private static byte[] CreateImagePdf(int pageCount) {
        PdfDocument document = PdfDocument.Create();
        for (int page = 1; page <= pageCount; page++) {
            if (page > 1) document.PageBreak();
            int number = page;
            document.Paragraph(paragraph => paragraph.AddText("Image page " + number));
            document.Image(CreateNoisyPng(900, 700, seed: page), 400, 300);
        }

        return document.ToBytes();
    }

    // Noise keeps each image near its raw size after compression, so the file is dominated by image payloads.
    private static byte[] CreateNoisyPng(int width, int height, int seed) {
        var random = new Random(seed);
        var scanlines = new byte[(1 + width * 3) * height];
        for (int y = 0; y < height; y++) {
            int row = y * (1 + width * 3);
            for (int x = 1; x <= width * 3; x++) scanlines[row + x] = (byte)random.Next(256);
        }

        using var compressed = new MemoryStream();
        using (var zlib = new System.IO.Compression.ZLibStream(compressed, System.IO.Compression.CompressionLevel.Fastest, leaveOpen: true)) {
            zlib.Write(scanlines);
        }

        using var png = new MemoryStream();
        png.Write(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 });
        var header = new byte[13];
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(header, width);
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(header.AsSpan(4), height);
        header[8] = 8;
        header[9] = 2;
        WriteChunk(png, "IHDR", header);
        WriteChunk(png, "IDAT", compressed.ToArray());
        WriteChunk(png, "IEND", Array.Empty<byte>());
        return png.ToArray();
    }

    private static void WriteChunk(Stream png, string type, byte[] data) {
        Span<byte> number = stackalloc byte[4];
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(number, data.Length);
        png.Write(number);
        byte[] typeBytes = Encoding.ASCII.GetBytes(type);
        png.Write(typeBytes);
        png.Write(data);
        uint crc = 0xFFFFFFFF;
        foreach (byte value in typeBytes.Concat(data)) {
            crc ^= value;
            for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? 0xEDB88320 ^ (crc >> 1) : crc >> 1;
        }
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32BigEndian(number, ~crc);
        png.Write(number);
    }
}
#endif
