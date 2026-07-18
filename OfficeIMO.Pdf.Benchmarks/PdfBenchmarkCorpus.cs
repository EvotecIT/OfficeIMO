using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

internal static class PdfBenchmarkCorpus {
    internal const int PageCount = 60;

    internal static byte[] Create() {
        byte[] image = CreateImage();
        PdfDocument document = PdfDocument.Create(new PdfOptions {
            CompressContentStreams = true,
            DefaultFontSize = 10
        }).Meta(title: "OfficeIMO.Pdf mixed performance corpus");

        for (int page = 1; page <= PageCount; page++) {
            document
                .H1("Operational report " + page)
                .Paragraph(paragraph => paragraph.Text(
                    "This deterministic page exercises parsing, text extraction, inspection, diagnostics, " +
                    "stream decoding, page-tree traversal, vector projection, image reuse, and managed rendering."))
                .Table(new[] {
                    new[] { "Metric", "Value", "Status" },
                    new[] { "Documents", (page * 37).ToString(), "Healthy" },
                    new[] { "Rules", (page * 11).ToString(), "Reviewed" },
                    new[] { "Signals", (page * 19).ToString(), "Observed" }
                })
                .Rectangle(
                    180,
                    24,
                    strokeColor: PdfColor.FromRgb(0, 64, 128),
                    strokeWidth: 1.5,
                    fillColor: PdfColor.FromRgb(220, 240, 252))
                .Image(image, 32, 32, alternativeText: "Deterministic benchmark image");
            if (page < PageCount) {
                document.PageBreak();
            }
        }

        byte[] bytes = document.ToBytes();
        int pages = PdfDocument.Open(bytes).Inspect().PageCount;
        if (pages != PageCount) {
            throw new InvalidOperationException($"Benchmark corpus produced {pages} pages instead of {PageCount}.");
        }

        return bytes;
    }

    private static byte[] CreateImage() {
        var image = new OfficeRasterImage(32, 32, OfficeColor.White);
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                image.SetPixel(
                    x,
                    y,
                    OfficeColor.FromRgb(
                        (byte)(32 + (x * 6)),
                        (byte)(48 + (y * 5)),
                        (byte)(160 - ((x + y) * 2))));
            }
        }

        return OfficePngWriter.Encode(image);
    }
}
