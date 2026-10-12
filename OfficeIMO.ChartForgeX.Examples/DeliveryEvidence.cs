using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Visio;
using OfficeIMO.Word;

namespace OfficeIMO.ChartForgeX.Examples;

internal static class DeliveryEvidence {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        WriteIndented = true, Converters = { new JsonStringEnumConverter() }
    };

    public static void RenderSavedDocuments(string output) {
        using (var document = WordDocument.Load(Path.Combine(output, "service-review.docx"))) {
            SaveImages(output, "word-layout", document.ExportImages(OfficeImageExportFormat.Png));
            SaveImages(output, "word-layout", document.ExportImages(OfficeImageExportFormat.Svg));
        }
        using (var document = ExcelDocument.Load(Path.Combine(output, "service-review.xlsx")))
            SaveImages(output, "excel-layout", document.ExportImages(OfficeImageExportFormat.Png,
                new ExcelWorkbookImageExportOptions { ShowGridlines = false }));
        using (var document = PowerPointPresentation.Load(Path.Combine(output, "service-review.pptx")))
            SaveImages(output, "powerpoint-layout", document.ExportImages(OfficeImageExportFormat.Png));
        var pdf = PdfDocument.Load(File.ReadAllBytes(Path.Combine(output, "service-review.pdf")));
        SaveImages(output, "saved-pdf", pdf.Render.ExportImages(OfficeImageExportFormat.Png));
        var visio = VisioDocument.Load(Path.Combine(output, "service-delivery.vsdx"));
        SaveImages(output, "visio-layout", visio.ExportImages(OfficeImageExportFormat.Png));
        SaveImages(output, "visio-layout", visio.ExportImages(OfficeImageExportFormat.Svg));
    }

    private static void SaveImages(string output, string route, IReadOnlyList<OfficeImageExportResult> images) {
        if (images.Count == 0) throw new InvalidOperationException(route + " did not render a saved document image.");
        var reports = new List<object>();
        for (int index = 0; index < images.Count; index++) {
            var image = images[index];
            string fileName = route + "-" + (index + 1).ToString("D2") +
                (image.Format == OfficeImageExportFormat.Svg ? ".svg" : ".png");
            File.WriteAllBytes(Path.Combine(output, fileName), image.Bytes);
            reports.Add(new { File = fileName, image.Width, image.Height, image.Diagnostics });
        }
        SaveJson(Path.Combine(output, route + "-" + images[0].Format.ToString().ToLowerInvariant() + ".json"), reports);
    }

    public static void SaveJson<T>(string path, T value) => File.WriteAllText(path, JsonSerializer.Serialize(value, JsonOptions));
}
