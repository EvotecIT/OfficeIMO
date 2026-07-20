using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public partial class ExcelImageExportTests {
    [Fact]
    public void ExcelRasterExport_ReachesTheSharedTextShapingProvider() {
        string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        using ExcelDocument document = ExcelDocument.Create(filePath);
        ExcelSheet sheet = document.AddWorksheet("Shaping");
        sheet.CellValue(1, 1, "A");
        sheet.CellAt(1, 1).SetFontName(ManagedTextShapingTestAssets.FamilyName);
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new ExcelImageExportOptions {
            ShowGridlines = false,
            TextShapingProvider = provider,
            TextShapingLanguage = "ar-SA"
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont('A'));

        OfficeImageExportResult result =
            sheet.Range("A1:A1").ExportImage(OfficeImageExportFormat.Png, options);

        Assert.Equal(OfficeImageExportFormat.Png, result.Format);
        Assert.Contains(provider.Requests, request =>
            request.Text == "A" &&
            request.Language == "ar-SA");
    }
}
