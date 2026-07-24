using OfficeIMO.Drawing;
using OfficeIMO.Drawing.HarfBuzz;
using OfficeIMO.Pdf;

internal static class PdfBenchmarkCorpus {
    internal const int PageCount = 60;
    private static readonly Lazy<byte[]> CarlitoRegular = new(
        static () => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "Carlito-Regular.ttf")),
        isThreadSafe: true);

    internal static byte[] Create() {
        return CreateDocument(PdfObjectSerializationMode.Buffered).ToBytes();
    }

    internal static PdfDocument CreateDocument(PdfObjectSerializationMode objectSerializationMode) {
        byte[] image = CreateImage();
        PdfDocument document = PdfDocument.Create(new PdfOptions {
            CompressContentStreams = true,
            DefaultFontSize = 10,
            FileVersion = objectSerializationMode == PdfObjectSerializationMode.ForwardOnly
                ? PdfFileVersion.Pdf17
                : PdfFileVersion.Pdf14,
            ObjectSerializationMode = objectSerializationMode
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

        return document;
    }

    internal static PdfDocument CreateHarfBuzzDocument() {
        byte[] fontData = CarlitoRegular.Value;
        var options = new PdfOptions {
            CompressContentStreams = true,
            DefaultFontSize = 10,
            FileVersion = PdfFileVersion.Pdf17,
            ObjectSerializationMode = PdfObjectSerializationMode.ForwardOnly,
            TextShapingMode = PdfTextShapingMode.LatinLigatures
        }.SetTextShapingProvider(OfficeHarfBuzzTextShapingProvider.Instance);
        options.RegisterFontFamily(
            PdfStandardFont.Helvetica,
            new PdfEmbeddedFontFamily("Carlito", fontData));
        PdfDocument document = PdfDocument.Create(options)
            .Meta(title: "OfficeIMO.Pdf HarfBuzz performance corpus");
        const string shapedText =
            "Office affinity efficient official workflow: AVATAR, ffi, fi, fl. " +
            "Repeated shaping must reuse one parsed font face across measurement and drawing.";
        for (int page = 1; page <= PageCount; page++) {
            document.H1("Shaped report " + page);
            for (int paragraph = 0; paragraph < 12; paragraph++) {
                document.Paragraph(builder => builder.Text(shapedText));
            }
            if (page < PageCount) {
                document.PageBreak();
            }
        }

        return document;
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
