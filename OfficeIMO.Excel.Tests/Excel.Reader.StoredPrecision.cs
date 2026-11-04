using System.Globalization;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Reader_stored_doubles_keep_precision_across_ranges_streams_and_formula_caches(bool customConverter) {
        double[] expected = { 0.6249999999999999d, -0.6249999999999999d, 0.12345678912345678d, double.Epsilon, 6.25e60d };
        string path = Path.Combine(_directoryWithFiles, "Reader.StoredDoublePrecision.xlsx");
        using (var document = ExcelDocument.Create(path)) {
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Value");
            sheet.Cell(1, 2).SetValue("Cache");
            sheet.Cell(1, 3).SetValue("Text");
            for (int index = 0; index < expected.Length; index++) {
                int row = index + 2;
                sheet.Cell(row, 1).SetValue(expected[index]);
                sheet.Cell(row, 2).SetValue(expected[index]);
                sheet.Cell(row, 2).SetFormula("=1");
                sheet.Cell(row, 3).SetValue(expected[index].ToString("R", CultureInfo.InvariantCulture));
            }
            sheet.Cell(4098, 1).SetValue(1d); // Include the supported large-range data-reader path.
            document.Save();
        }
        var options = new ExcelReaderOptions();
        if (customConverter) options.CellValueConverter = _ => ExcelCellValue.NotHandled;
        using var reader = ExcelDocumentReader.Open(path, options);
        var sheetReader = reader.GetSheet("Data");
        object?[,] range = sheetReader.ReadRange("A2:C6", ExcelExecutionMode.Sequential);
        var streamed = sheetReader.ReadRangeStream("A2:C6", chunkRows: 1, mode: ExcelExecutionMode.Sequential).ToArray();
        var objects = sheetReader.ReadObjects<StoredPrecisionRow>("A1:C6").ToArray();
        var column = sheetReader.ReadColumn("A2:A6").ToArray();
        using var dataReader = sheetReader.ReadRangeAsDataReader("A1:C4098", schemaSampleRows: 0);
        for (int index = 0; index < expected.Length; index++) {
            Assert.Equal(expected[index], Assert.IsType<double>(range[index, 0]));
            Assert.Equal(expected[index], Assert.IsType<double>(range[index, 1]));
            Assert.Equal(expected[index], Assert.IsType<double>(streamed[index].Rows[0][0]));
            Assert.Equal(expected[index], Assert.IsType<double>(streamed[index].Rows[0][1]));
            Assert.Equal(expected[index], objects[index].Value);
            Assert.Equal(expected[index], objects[index].Cache);
            Assert.Equal(expected[index], Assert.IsType<double>(column[index]));
            Assert.True(dataReader.Read());
            Assert.Equal(expected[index], Assert.IsType<double>(dataReader.GetValue(0)));
            Assert.Equal(expected[index], Assert.IsType<double>(dataReader.GetValue(1)));
            string text = expected[index].ToString("R", CultureInfo.InvariantCulture);
            Assert.Equal(text, range[index, 2]);
            Assert.Equal(text, objects[index].Text);
            Assert.Equal(text, dataReader.GetString(2));
        }
    }

    [Fact]
    public void SaveAsPdf_saved_fraction_values_match_image_snapshot_across_midpoints_and_large_integers() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Fraction precision");
        sheet.Cell(1, 1).SetValue(0.6249999999999999d).SetNumberFormat("# ?/4;-# ?/4;0");
        sheet.Cell(2, 1).SetValue(-0.6249999999999999d);
        sheet.Cell(2, 1).SetFormula("=1");
        sheet.Cell(2, 1).SetNumberFormat("# ?/4;-# ?/4;0");
        sheet.Cell(3, 1).SetValue(1e20d).SetNumberFormat("# ?/8;-# ?/8;0");
        sheet.Cell(4, 1).SetValue(-1e20d).SetNumberFormat("# ?/8;-# ?/8;0");
        // Wide enough that Excel would show the 21-digit values rather than "####".
        sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(30));
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(new[] { "2/4", "-2/4", "100000000000000000000", "-100000000000000000000" },
            reopened.Sheets[0].Range("A1:A4").CreateVisualSnapshot().Cells.Select(c => c.Text));
        byte[] bytes = reopened.ToPdf(new ExcelPdfExportOptions { IncludeSheetHeadings = false, HeaderRowCount = 0 }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using var pdf = PdfPigDocument.Open(bytes);
        string text = string.Concat(Enumerable.Range(1, pdf.NumberOfPages).Select(page => pdf.GetPage(page).Text));
        Assert.Contains("2/4", text);
        Assert.Contains("-2/4", text);
        Assert.DoesNotContain("3/4", text);
        Assert.Contains("100000000000000000000", text);
        Assert.Contains("-100000000000000000000", text);
        Assert.DoesNotContain("/8", text);
    }

    private sealed class StoredPrecisionRow {
        public double Value { get; set; }
        public double Cache { get; set; }
        public string? Text { get; set; }
    }
}
