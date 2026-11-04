using System.IO.Compression;
using System.Management.Automation;
using System.Xml.Linq;
using OfficeIMO.Excel;

namespace PSWriteOffice.Tests;

public sealed partial class ConversionCmdletWorkflowTests {
    [Theory]
    [InlineData("Set-OfficeExcelHeader")]
    [InlineData("Set-OfficeExcelFooter")]
    public void ExcelRunningTextCommandsPreserveTheOtherHalfAndBindPhysicalImageDimensionsAndFalseSwitches(string command) {
        string root = Directory.CreateTempSubdirectory("officeimo-excel-running-text-").FullName;
        try {
            string path = Path.Combine(root, "logo[1].png");
            byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
            File.WriteAllBytes(path, png);
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Report");
            sheet.PageSetup.Header.Left.Text = "Header"; sheet.PageSetup.Footer.Left.Text = "Footer";
            sheet.PageSetup.DifferentFirstPage = true; sheet.PageSetup.DifferentOddEven = true;
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("document", document);
            runspace.SessionStateProxy.SetVariable("sheet", sheet);
            runspace.SessionStateProxy.SetVariable("path", path);
            runspace.SessionStateProxy.SetVariable("command", command);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("""
                & $command -Document $document -Sheet $sheet -Center 'Review' -ImagePath $path `
                    -ImagePosition Center -ImageWidth '2.54cm' -ImageHeight '12.7mm' `
                    -NoAlignWithMargins -NoScaleWithDocument
                & $command -Document $document -Sheet 1 -Right 'Page &P of &N' `
                    -DifferentFirstPage:$false -NoAlignWithMargins:$false -PassThru
                """);
            var output = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.IsType<ExcelSheet>(Assert.Single(output).BaseObject);
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using (var reopened = ExcelDocument.Load(saved)) {
                var page = reopened.Sheets[0].PageSetup;
                var edited = command.EndsWith("Header", StringComparison.Ordinal) ? page.Header : page.Footer;
                Assert.Equal("Header", page.Header.Left.Text); Assert.Equal("Footer", page.Footer.Left.Text);
                Assert.Contains("Review", edited.Center.Text, StringComparison.Ordinal);
                Assert.False(page.DifferentFirstPage); Assert.True(page.DifferentOddEven);
                var image = edited.Center.Image;
                Assert.NotNull(image);
                Assert.Equal(png, image.Bytes);
                Assert.Equal("image/png", image.ContentType);
                Assert.Equal(72, image.Width.ToPoints(), 3); Assert.Equal(36, image.Height.ToPoints(), 3);
                Assert.Equal("Page &P of &N", edited.Right.Text);
            }
            saved.Position = 0;
            using var zip = new ZipArchive(saved, ZipArchiveMode.Read);
            using var xmlStream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var runningText = XDocument.Load(xmlStream).Root!.Element(ns + "headerFooter")!;
            Assert.Equal("1", (string?) runningText.Attribute("alignWithMargins"));
            Assert.Equal("0", (string?) runningText.Attribute("scaleWithDoc"));
        } finally { Directory.Delete(root, true); }
    }
}
