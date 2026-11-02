using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class ConversionCmdletWorkflowTests {
    [Fact]
    public void NativeImageNestedOverridesRetainTypedPageSettingsAndCallerOptions() {
        using var word = OfficeIMO.Word.WordDocument.Create();
        word.AddParagraph("Nested option evidence");
        var options = new OfficeIMO.Word.WordImageExportOptions {
            LayoutOptions = new OfficeIMO.Word.Pdf.WordPdfExportOptions {
                PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeLength.Points(400), OfficeLength.Points(500))
            }
        };
        using var runspace = OpenRunspace();
        runspace.SessionStateProxy.SetVariable("document", word);
        runspace.SessionStateProxy.SetVariable("options", options);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("$document | ConvertTo-OfficeDocumentImage -Format Svg -Options @($options, @{ LayoutOptions = @{ TrackedChanges = 'Final' } }) -ErrorAction Stop");
        var result = shell.Invoke();
        Assert.Empty(shell.Streams.Error);
        var image = Assert.IsType<OfficeImageExportResult>(Assert.Single(result).BaseObject);
        Assert.Equal(400, image.Width);
        Assert.Equal(500, image.Height);
        Assert.Equal(OfficeIMO.Word.Pdf.WordPdfTrackedChangeView.Markup, options.LayoutOptions.TrackedChanges);
        Assert.Equal(400D, options.LayoutOptions.PageSize!.Value.Width.ToPoints());
    }

    [Fact]
    public void NativeImagePowerPointSelectionsKeepLiveEditsReadOnlySourcesAndCallerOptions() {
        string root = Directory.CreateTempSubdirectory("officeimo-native-image-slides-").FullName;
        using var document = PowerPointDocument.Create();
        document.SlideSize.SetSize(OfficeLength.Points(120), OfficeLength.Points(80));
        foreach (string color in new[] { "112233", "445566", "778899" }) document.AddSlide().BackgroundColor = color;
        const string password = "synthetic-image-password";
        try {
            string source = Path.Combine(root, "deck[1].pptx"), outputPath = Path.Combine(root, "selected.svg");
            byte[] saved = document.ToBytes(new PowerPointSaveOptions { Password = password });
            File.WriteAllBytes(source, saved);
            using var readOnly = PowerPointDocument.Load(source, new PowerPointLoadOptions { Password = password, AccessMode = DocumentAccessMode.ReadOnly });
            var retainedSlide = document.Slides[1];
            retainedSlide.BackgroundColor = "AABBCC";
            var options = new PowerPointImageExportOptions { SlideNumbers = new[] { 2 }, Scale = 1.5D };
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("live", document);
            runspace.SessionStateProxy.SetVariable("readOnly", readOnly);
            runspace.SessionStateProxy.SetVariable("options", options);
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("outputPath", outputPath);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("""
                $live | ConvertTo-OfficeDocumentImage -Page '3,1,3' -Format Svg -Options $options -ErrorAction Stop
                $live | Export-OfficeDocumentImage -Page 2 -OutputPath $outputPath -ErrorAction Stop
                $readOnly | ConvertTo-OfficeDocumentImage -Page 2 -Format Svg -ErrorAction Stop
                [pscustomobject]@{ FullName = $source } | ConvertTo-OfficeDocumentImage -Page 2 -Format Svg -Password synthetic-image-password -ErrorAction Stop
                """);
            var result = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            Assert.Equal(5, result.Count);
            var images = result.Select(item => Assert.IsType<OfficeImageExportResult>(item.BaseObject)).ToArray();
            Assert.Equal(new[] { "Slide 1", "Slide 3" }, images.Take(2).Select(image => image.Name));
            Assert.Contains("#AABBCC", System.Text.Encoding.UTF8.GetString(images[2].Bytes), StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#445566", System.Text.Encoding.UTF8.GetString(images[3].Bytes), StringComparison.OrdinalIgnoreCase);
            Assert.Equal(images[3].Bytes, images[4].Bytes);
            Assert.Equal(outputPath, images[2].SavedPath);
            Assert.Equal(new[] { 2 }, options.SlideNumbers);
            Assert.Equal(1.5D, options.Scale);
            Assert.Equal(saved, File.ReadAllBytes(source));
            Assert.Same(retainedSlide, document.Slides[1]);
            Assert.Equal(DocumentAccessMode.ReadOnly, readOnly.AccessMode);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void NativeImageExcelSheetRangeAndChartSelectorsUseNativeExportSettings() {
        string root = Directory.CreateTempSubdirectory("officeimo-native-image-sheets-").FullName;
        using var book = ExcelDocument.Create();
        var sheet = book.AddSheet("Dashboard");
        sheet.CellValue(1, 1, "Region"); sheet.CellValue(1, 2, "Revenue");
        sheet.CellValue(2, 1, "East"); sheet.CellValue(2, 2, 3);
        sheet.CellValue(3, 1, "West"); sheet.CellValue(3, 2, 6);
        book.AddSheet("Other").CellValue(1, 1, "Other sheet");
        var chart = sheet.AddChart(OfficeChartType.BarClustered, "A1:B3", new ExcelChartOptions { Row = 2, Column = 5 });
        chart.Name = "Revenue";
        try {
            string outputPath = Path.Combine(root, "chart.svg"), invalidPath = Path.Combine(root, "new", "multiple.svg");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("book", book); runspace.SessionStateProxy.SetVariable("sheet", sheet);
            runspace.SessionStateProxy.SetVariable("outputPath", outputPath); runspace.SessionStateProxy.SetVariable("invalidPath", invalidPath);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("""
                $book | ConvertTo-OfficeDocumentImage -Sheet $sheet -Range A1:B3 -Format Svg -Options @{ ShowGridlines = $false } -ErrorAction Stop
                $book | Export-OfficeDocumentImage -Sheet Dashboard -Chart Revenue -OutputPath $outputPath -Dpi 144 -ErrorAction Stop
                $refused = $false
                try { $book | Export-OfficeDocumentImage -Sheet 1,2 -Range A1:B3 -OutputPath $invalidPath -ErrorAction Stop }
                catch { $refused = $true }
                if (-not $refused) { throw 'Multiple sheets were saved to one image file' }
                """);
            var result = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            Assert.Equal(2, result.Count);
            var range = Assert.IsType<OfficeImageExportResult>(result[0].BaseObject);
            var actualChart = Assert.IsType<OfficeImageExportResult>(result[1].BaseObject);
            Assert.Equal(sheet.ExportImage(OfficeImageExportFormat.Svg, new ExcelImageExportOptions { Range = "A1:B3", ShowGridlines = false }).Bytes, range.Bytes);
            Assert.Equal(chart.ExportImage(OfficeImageExportFormat.Svg, new ExcelImageExportOptions { TargetDpi = 144 }).Bytes, actualChart.Bytes);
            Assert.Equal(outputPath, actualChart.SavedPath);
            Assert.False(Directory.Exists(Path.GetDirectoryName(invalidPath)));
            Assert.Equal("East", sheet.CellAt(2, 1).GetValue<string>());
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void NativeImageFileBatchNoClobberRetainsTheFirstSharedFilenameAndReportsTheCollision() {
        string root = Directory.CreateTempSubdirectory("officeimo-native-image-collision-").FullName;
        try {
            string first = Path.Combine(root, "first", "report.html"), second = Path.Combine(root, "second", "report.html");
            Directory.CreateDirectory(Path.GetDirectoryName(first)!); Directory.CreateDirectory(Path.GetDirectoryName(second)!);
            File.WriteAllText(first, "<h1 style='color:#112233'>First report</h1>");
            File.WriteAllText(second, "<h1 style='color:#AABBCC'>Second report</h1>");
            string outputDirectory = Path.Combine(root, "images");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("inputs", new[] { first, second });
            runspace.SessionStateProxy.SetVariable("outputDirectory", outputDirectory);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("""
                Export-OfficeDocumentImage -LiteralPath $inputs -OutputDirectory $outputDirectory -Format Svg -NoClobber -ThrottleLimit 1 -ErrorAction SilentlyContinue -ErrorVariable collision
                $collision.Count
                """);
            var result = shell.Invoke();
            Assert.Equal(2, result.Count);
            var image = Assert.IsType<OfficeImageExportResult>(result[0].BaseObject);
            Assert.Equal(1, result[1].BaseObject);
            Assert.Equal(Path.Combine(outputDirectory, "report-1.svg"), image.SavedPath);
            Assert.Equal(image.Bytes, File.ReadAllBytes(image.SavedPath!));
            Assert.Contains("#112233", File.ReadAllText(image.SavedPath!), StringComparison.OrdinalIgnoreCase);
            Assert.Single(Directory.GetFiles(outputDirectory));
        } finally { Directory.Delete(root, true); }
    }
}
