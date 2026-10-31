using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Write;
using System.IO.Compression;
using System.Threading;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("general", false)]
        [InlineData("general", true)]
        [InlineData("direct", false)]
        [InlineData("direct", true)]
        [InlineData("staged", false)]
        [InlineData("staged", true)]
        public void Xlsb_NewWorkbookRetainsItsDrawingMlThemeOnEveryWriterPath(string route, bool sharedStrings) {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddSheet("Data").CellValue(1, 1, "Theme evidence");
            document.ResetWorkbookTheme("Custom & bounded theme");
            XDocument source = XDocument.Parse(document.GetWorkbookThemeXml()!);
            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            source.Descendants(drawing + "accent1").Single().Elements().Single().SetAttributeValue("val", "123ABC");
            document.SetWorkbookThemeXml(source.ToString());
            using var destination = new MemoryStream();
            if (route == "general") {
                document.AddSheet("Other").CellValue(1, 1, "Second sheet");
                document.Save(destination, ExcelFileFormat.Xlsb, new ExcelSaveOptions { XlsbUseSharedStrings = sharedStrings });
            } else {
                var rows = new SnapshotXlsbRows(aboveCaptureLimit: route == "staged");
                var tabular = new ExcelDirectTabularSource("Data", rows, includeHeaders: true, preserveMissingValues: true);
                Assert.True(XlsbNewPackageWriter.TryWriteDirectTabular(document, tabular, destination, CancellationToken.None, sharedStrings));
            }
            byte[] bytes = destination.ToArray();
            XDocument theme = XDocument.Parse(System.Text.Encoding.UTF8.GetString(ReadPackageEntry(bytes, "xl/theme/theme1.xml")));
            Assert.True(XNode.DeepEquals(source, theme));
            Assert.Empty(new OpenXmlValidator().Validate(document.WorkbookPartRoot.ThemePart!.Theme!));
            XNamespace package = "http://schemas.openxmlformats.org/package/2006/relationships";
            XDocument relationships = XDocument.Parse(System.Text.Encoding.UTF8.GetString(ReadPackageEntry(bytes, "xl/_rels/workbook.bin.rels")));
            XElement link = Assert.Single(relationships.Root!.Elements(package + "Relationship"), item =>
                (string?)item.Attribute("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme");
            Assert.Equal("theme/theme1.xml", (string?)link.Attribute("Target"));
            Assert.Equal(relationships.Root!.Elements().Count(), relationships.Root.Elements().Select(item => (string?)item.Attribute("Id")).Distinct().Count());
            XDocument contentTypes = XDocument.Parse(System.Text.Encoding.UTF8.GetString(ReadPackageEntry(bytes, "[Content_Types].xml")));
            Assert.Contains(contentTypes.Root!.Elements(), item => (string?)item.Attribute("PartName") == "/xl/theme/theme1.xml"
                && (string?)item.Attribute("ContentType") == "application/vnd.openxmlformats-officedocument.theme+xml");
            using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(bytes, writable: false));
            Assert.True(XNode.DeepEquals(theme, XDocument.Parse(loaded.GetWorkbookThemeXml()!)));
        }

        [Fact]
        public void Xlsb_ExcelProducedThemeIsProjectedAndPreservedDuringNativeCellEdits() {
            string path = Path.Combine(AppContext.BaseDirectory, "Documents", "XlsbCorpus", "excel-generated", "styles-dates-formulas.xlsb");
            byte[] original = File.ReadAllBytes(path);
            byte[] theme = ReadPackageEntry(original, "xl/theme/theme1.xml");
            using ExcelDocument document = ExcelDocument.Load(new MemoryStream(original, writable: false));
            Assert.True(XNode.DeepEquals(XDocument.Parse(System.Text.Encoding.UTF8.GetString(theme)), XDocument.Parse(document.GetWorkbookThemeXml()!)));
            document.Sheets[0].CellValue(1, 1, "Edited theme evidence");
            byte[] edited = document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
            Assert.Equal(theme, ReadPackageEntry(edited, "xl/theme/theme1.xml"));
        }

        [Fact]
        public void Xlsb_NewWriterRefusesToDropRelatedThemePartsBeforeWriting() {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddSheet("Data").CellValue(1, 1, "Theme evidence");
            document.EnsureWorkbookTheme();
            document.WorkbookPartRoot.ThemePart!.AddExternalRelationship("urn:unsupported-theme-carrier", new Uri("https://example.invalid/inert"));
            using var destination = new MemoryStream();
            destination.WriteByte(42);
            Assert.Throws<NotSupportedException>(() => XlsbNewPackageWriter.Write(document, destination));
            Assert.Equal(new byte[] { 42 }, destination.ToArray());
        }
    }
}
