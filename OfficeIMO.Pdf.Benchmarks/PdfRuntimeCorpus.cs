using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks;

/// <summary>Deterministic small business pages for opt-in engine measurements.</summary>
internal static class PdfRuntimeCorpus {
    internal const double PageWidth = 600D;
    internal const double PageHeight = 800D;
    internal const string BusinessProfile = "BusinessEmbeddedCarlito";
    internal const string EditProfile = "TextEditStandard14";
    private static readonly byte[] Image = CreateImage();
    private static readonly Lazy<byte[]> CarlitoRegular = new(
        static () => File.ReadAllBytes(Path.Combine(Path.GetDirectoryName(typeof(PdfRuntimeWorkload).Assembly.Location)!, "Fonts", "Carlito-Regular.ttf")));

    internal static string Title(int page) => "Runtime report " + Number(page);
    internal static string Target(int page) => "Editable marker" + Number(page);
    internal static string Replacement(int page) => "Edited marker" + Number(page);
    internal static string Neighbor(int page) => "Retained neighbor " + Number(page);
    internal static string Number(int page) => page.ToString("0000", CultureInfo.InvariantCulture);

    internal static IReadOnlyList<string> ExpectedText(int page, bool replaced, bool textOnly) => textOnly ? new[] {
        Title(page), Neighbor(page), replaced ? Replacement(page) : Target(page), "Validated page " + Number(page)
    } : new[] {
        Title(page), Neighbor(page), replaced ? Replacement(page) : Target(page),
        "Metric", "Value", "Status", "Documents", (page * 37).ToString(CultureInfo.InvariantCulture),
        "Healthy", "Amount", (page * 101.25D).ToString("0.00", CultureInfo.InvariantCulture),
        "Reviewed", "Validated page " + Number(page)
    };

    internal static PdfDocument Create(int pages, PdfObjectSerializationMode mode, bool textOnly = false) {
        var options = new PdfCreateOptions {
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12),
            CompressContentStreams = true,
            FileVersion = mode == PdfObjectSerializationMode.ForwardOnly ? PdfFileVersion.Pdf17 : PdfFileVersion.Pdf14,
            ObjectSerializationMode = mode
        };
        if (!textOnly) options.RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Carlito", CarlitoRegular.Value));
        return PdfDocument.Create(compose => {
            for (int page = 1; page <= pages; page++) {
                int number = page;
                compose.Page(layout => layout.Size(PageWidth, PageHeight).Content(content => {
                    content.Paragraph(text => text.Text(Title(number)))
                        .Paragraph(text => text.Text(Neighbor(number)))
                        .Paragraph(text => text.Text(Target(number)));
                    if (!textOnly) content.Table(new[] {
                        new[] { "Metric", "Value", "Status" },
                        new[] { "Documents", (number * 37).ToString(CultureInfo.InvariantCulture), "Healthy" },
                        new[] { "Amount", (number * 101.25D).ToString("0.00", CultureInfo.InvariantCulture), "Reviewed" }
                    });
                    content.Paragraph(text => text.Text("Validated page " + Number(number)));
                    if (!textOnly) content.Rectangle(global::OfficeIMO.OfficeLength.Points(180), global::OfficeIMO.OfficeLength.Points(24), fillColor: global::OfficeIMO.OfficeColor.FromRgb(220, 240, 252))
                        .Image(Image, global::OfficeIMO.OfficeLength.Points(32), global::OfficeIMO.OfficeLength.Points(32), alternativeText: "Deterministic runtime measurement image");
                }));
            }
        }, options).Meta(title: "OfficeIMO PDF runtime corpus");
    }

    private static byte[] CreateImage() {
        var image = new OfficeRasterImage(32, 32, OfficeColor.White);
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                image.SetPixel(x, y, OfficeColor.FromRgb((byte)(32 + x * 6), (byte)(48 + y * 5), (byte)(160 - (x + y) * 2)));
            }
        }
        return OfficePngWriter.Encode(image);
    }
}
