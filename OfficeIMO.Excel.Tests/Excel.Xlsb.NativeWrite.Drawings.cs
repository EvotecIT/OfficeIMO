using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using System.IO.Compression;
using System.Text;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

/// <summary>Worksheet images and charts written by the native XLSB writers and projected by the XLSB reader.</summary>
public sealed class ExcelXlsbDrawingTests {
    [Fact]
    public void Images_keep_their_bytes_and_anchor_in_native_xlsb() {
        byte[] png = Png();
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Logo");
            sheet.AddImage(3, 2, png, new ExcelImageOptions {
                ContentType = "image/png", Width = OfficeLength.Pixels(64), Height = OfficeLength.Pixels(32)
            }).Name = "Logo";
        });

        byte[] xlsb = SaveAsXlsb(xlsx);

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(xlsb));
        ExcelImage image = Assert.Single(reloaded["Data"].Images);
        Assert.Equal(png, image.ToBytes());
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal("Logo", image.Name);
        Assert.Equal(3, image.RowIndex);
        Assert.Equal(2, image.ColumnIndex);
        Assert.Equal(OfficeLength.Pixels(64), image.Width);
        Assert.Equal(OfficeLength.Pixels(32), image.Height);
    }

    [Fact]
    public void Native_drawing_cannot_take_the_name_of_a_projected_worksheet() {
        byte[] png = Png();
        byte[] source = SaveAsXlsb(CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Images");
            sheet.Cell(1, 1).SetValue("Logo");
            sheet.AddImage(3, 2, png, new ExcelImageOptions { ContentType = "image/png" });
            document.AddSheet("Other").Cell(1, 1).SetValue(42);
        }));
        using var modified = new MemoryStream();
        using (var input = new ZipArchive(new MemoryStream(source), ZipArchiveMode.Read))
        using (var output = new ZipArchive(modified, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ZipArchiveEntry entry in input.Entries) {
                string name = entry.FullName switch {
                    "xl/drawings/drawing1.xml" => "xl/worksheets/sheet2.xml",
                    "xl/drawings/_rels/drawing1.xml.rels" => "xl/worksheets/_rels/sheet2.xml.rels",
                    _ => entry.FullName
                };
                using Stream part = entry.Open();
                using Stream replacement = output.CreateEntry(name).Open();
                if (name is "[Content_Types].xml" or "xl/worksheets/_rels/sheet1.bin.rels") {
                    using var reader = new StreamReader(part);
                    string oldTarget = name == "[Content_Types].xml" ? "/xl/drawings/drawing1.xml" : "../drawings/drawing1.xml";
                    string newTarget = name == "[Content_Types].xml" ? "/xl/worksheets/sheet2.xml" : "sheet2.xml";
                    string xml = reader.ReadToEnd();
                    Assert.Contains(oldTarget, xml, StringComparison.Ordinal);
                    byte[] bytes = Encoding.UTF8.GetBytes(xml.Replace(oldTarget, newTarget, StringComparison.Ordinal));
                    replacement.Write(bytes, 0, bytes.Length);
                } else {
                    part.CopyTo(replacement);
                }
            }
        }

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(modified.ToArray()));
        Assert.True(reloaded.IsSdkDeferred);
        Assert.Equal(42, reloaded["Other"].Cell(1, 1).GetValue<int>());
        Assert.Equal(png, Assert.Single(reloaded["Images"].Images).ToBytes());
    }

    [Theory]
    [InlineData(OfficeChartType.Column, false)]
    [InlineData(OfficeChartType.Line, false)]
    [InlineData(OfficeChartType.Pie, true)]
    public void Charts_keep_their_type_and_series_references_in_native_xlsb(OfficeChartType type, bool styled) {
        string[] expectedFormulas = Array.Empty<string>();
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Sales Data");
            WriteSales(sheet);
            ExcelChart created = sheet.AddChart(type, "A1:C4", new ExcelChartOptions { Row = 6, Column = 2, Title = "Sales" });
            if (styled) created.ApplyStylePreset();
            expectedFormulas = SeriesFormulas(sheet);
        });

        byte[] xlsb = SaveAsXlsb(xlsx);

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(xlsb));
        ExcelSheet data = reloaded["Sales Data"];
        ExcelChart chart = Assert.Single(data.Charts);
        Assert.Equal(type, chart.ChartType);
        Assert.Equal("Sales", chart.Title);
        Assert.Contains("'Sales Data'!$A$2:$A$4", expectedFormulas);
        Assert.Equal(expectedFormulas, SeriesFormulas(data));
        Assert.True(chart.TryGetData(out ExcelChartData chartData));
        Assert.Equal(new[] { "North", "South", "West" }, chartData.Categories);
        Assert.Equal(new[] { 10D, 20D, 30D }, chartData.Series[0].Values);
        ChartPart part = Assert.Single(data.WorksheetPart.DrawingsPart!.ChartParts);
        Assert.Equal(styled, part.GetPartsOfType<ChartStylePart>().Any());
        Assert.Equal(styled, part.GetPartsOfType<ChartColorStylePart>().Any());
    }

    [Fact]
    public void Editing_and_renaming_a_loaded_xlsb_keeps_its_drawing_parts() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            WriteSales(sheet);
            sheet.AddImage(8, 1, Png(), new ExcelImageOptions { ContentType = "image/png" });
            sheet.AddChart(OfficeChartType.Column, "A1:C4", new ExcelChartOptions { Row = 6, Column = 4 });
        });
        byte[] source = SaveAsXlsb(xlsx);

        byte[] unchanged;
        byte[] edited;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            unchanged = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
            loaded["Data"].Cell(2, 2).SetValue(15);
            loaded["Data"].Name = "Values";
            edited = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }

        Assert.Equal(source, unchanged);
        // The drawing and image parts are copied; only the chart's sheet references follow the rename.
        string drawing = Assert.Single(EntryNames(source), name => name.StartsWith("xl/drawings/drawing", StringComparison.Ordinal));
        string media = Assert.Single(EntryNames(source), name => name.StartsWith("xl/media/", StringComparison.Ordinal));
        Assert.Equal(ReadEntry(source, drawing), ReadEntry(edited, drawing));
        Assert.Equal(ReadEntry(source, media), ReadEntry(edited, media));
        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(edited));
        ExcelSheet values = reloaded["Values"];
        Assert.Single(values.Images);
        Assert.Equal(OfficeChartType.Column, Assert.Single(values.Charts).ChartType);
        Assert.Contains("'Values'!$B$2:$B$4", SeriesFormulas(values));
        Assert.Equal(15D, values.Cell(2, 2).GetValue<double>());
    }

    [Fact]
    public void Renaming_after_saving_a_new_workbook_as_xlsb_updates_its_chart_references() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsb");
        try {
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddSheet("Data");
                WriteSales(sheet);
                sheet.AddChart(OfficeChartType.Column, "A1:C4", new ExcelChartOptions { Row = 6, Column = 2 });
                document.Save(path);
                sheet.Name = "Values";
                document.Save();
            }

            using ExcelDocument reloaded = ExcelDocument.Load(path);
            Assert.Contains("'Values'!$B$2:$B$4", SeriesFormulas(reloaded["Values"]));
            Assert.DoesNotContain(SeriesFormulas(reloaded["Values"]), formula => formula.IndexOf("Data", StringComparison.Ordinal) >= 0);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void ChartEx_charts_are_refused_by_the_native_xlsb_writer() {
        using ExcelDocument document = ExcelDocument.Create();
        document.AddSheet("Funnel").AddModernChart(new ExcelChartData(new[] { "Leads", "Deals" },
            new[] { new ExcelChartSeries("Count", new[] { 100D, 10D }) }), 2, 2, ExcelModernChartType.Funnel);

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb }));
        Assert.Contains("chartEx", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Adding_an_image_to_a_loaded_xlsb_is_refused() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => document.AddSheet("Data").Cell(1, 1).SetValue("Plain")));

        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source));
        loaded["Data"].AddImage(2, 2, Png(), new ExcelImageOptions { ContentType = "image/png" });

        Assert.Throws<NotSupportedException>(() => loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb }));
    }

    private static void WriteSales(ExcelSheet sheet) {
        sheet.Cell(1, 1).SetValue("Region");
        sheet.Cell(1, 2).SetValue("Q1");
        sheet.Cell(1, 3).SetValue("Q2");
        string[] regions = { "North", "South", "West" };
        for (int index = 0; index < regions.Length; index++) {
            sheet.Cell(index + 2, 1).SetValue(regions[index]);
            sheet.Cell(index + 2, 2).SetValue((index + 1) * 10D);
            sheet.Cell(index + 2, 3).SetValue((index + 1) * 7D);
        }
    }

    private static string[] SeriesFormulas(ExcelSheet sheet) =>
        sheet.WorksheetPart.DrawingsPart!.ChartParts
            .SelectMany(part => part.ChartSpace!.Descendants<C.Formula>())
            .Select(formula => formula.Text)
            .ToArray();

    private static byte[] Png() {
        var image = new OfficeRasterImage(4, 2, OfficeColor.Transparent);
        image.Fill(OfficeColor.Blue);
        return OfficePngWriter.Encode(image);
    }

    private static byte[] CreateXlsx(Action<ExcelDocument> build) {
        using ExcelDocument document = ExcelDocument.Create();
        build(document);
        return document.ToBytes();
    }

    private static byte[] SaveAsXlsb(byte[] xlsx) {
        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(xlsx));
        return loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
    }

    private static string[] EntryNames(byte[] package) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        return archive.Entries.Select(entry => entry.FullName).ToArray();
    }

    private static byte[] ReadEntry(byte[] package, string name) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var output = new MemoryStream();
        using (Stream entry = archive.GetEntry(name)!.Open()) entry.CopyTo(output);
        return output.ToArray();
    }
}
