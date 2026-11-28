using System.Data;
using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Tests.TestStreams;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void WorksheetViewIgnoresShadowViewsInExtensionsAndUnsupportedAlternatives() {
            string path = Path.Combine(_directoryWithFiles, "WorksheetViewExtensions.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                document.AddSheet("Data").Cell(1, 1).SetValue("A");
                document.Save();
            }

            XNamespace sheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace compatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";
            const string partName = "xl/worksheets/sheet1.xml";
            XElement ShadowView() => new XElement(sheet + "sheetViews",
                new XElement(sheet + "sheetView", new XAttribute("workbookViewId", 0),
                    new XAttribute("showGridLines", 0), new XAttribute("zoomScale", 75),
                    new XElement(sheet + "pane", new XAttribute("state", "frozen"),
                        new XAttribute("xSplit", 2), new XAttribute("ySplit", 3))));
            var extensions = new XElement(sheet + "extLst",
                new XElement(sheet + "ext", new XAttribute("uri", "urn:officeimo:test:shadow-view"), ShadowView()));

            using (var archive = ZipFile.Open(path, ZipArchiveMode.Update)) {
                ZipArchiveEntry part = archive.GetEntry(partName)!;
                XDocument worksheet;
                using (Stream input = part.Open()) worksheet = XDocument.Load(input);
                XElement root = worksheet.Root!;
                root.Elements(sheet + "sheetViews").Remove();
                if (root.Element(sheet + "dimension") == null) {
                    root.AddFirst(new XElement(sheet + "dimension", new XAttribute("ref", "A1")));
                }
                if (root.Element(sheet + "sheetPr") == null) root.AddFirst(new XElement(sheet + "sheetPr"));
                root.SetAttributeValue(XNamespace.Xmlns + "future", "urn:officeimo:test:unsupported-view");
                root.Add(new XElement(compatibility + "AlternateContent",
                    new XElement(compatibility + "Choice", new XAttribute("Requires", "future"), ShadowView()),
                    new XElement(compatibility + "Fallback", ShadowView())));
                root.Add(extensions);
                part.Delete();
                using Stream output = archive.CreateEntry(partName).Open();
                worksheet.Save(output);
            }

            var loadOptions = new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly };
            using (var document = ExcelDocument.Load(path, loadOptions)) {
                Assert.True(document.IsSdkDeferred);
                ExcelSheetView view = document["Data"].View;
                ExcelWorksheetViewInfo inspected = view.Inspect();
                Assert.False(inspected.HasPane);
                Assert.True(view.ShowGridlines);
                Assert.Null(view.ZoomScale);
                Assert.Empty(document.OpenXmlSdkFallbacks);
            }

            using (var document = ExcelDocument.Load(path)) {
                Assert.True(document.IsSdkDeferred);
                ExcelSheetView view = document["Data"].View;
                view.ShowGridlines = false;
                view.Freeze(1, 1);
                document.Save();
                Assert.Empty(document.OpenXmlSdkFallbacks);
            }

            using var edited = ZipFile.OpenRead(path);
            using Stream stream = edited.GetEntry(partName)!.Open();
            XElement outputRoot = XDocument.Load(stream).Root!;
            XElement liveView = Assert.Single(outputRoot.Elements(sheet + "sheetViews"));
            Assert.Equal(new[] { sheet + "sheetPr", sheet + "dimension", sheet + "sheetViews" },
                outputRoot.Elements().Take(3).Select(element => element.Name));
            Assert.Equal("1", (string?)liveView.Element(sheet + "sheetView")?.Element(sheet + "pane")?.Attribute("ySplit"));
            Assert.Equal(sheet + "extLst", outputRoot.Elements().Last().Name);
            XElement writtenExtensions = outputRoot.Element(sheet + "extLst")!;
            Assert.True(XNode.DeepEquals(WithoutNamespaceDeclarations(extensions), WithoutNamespaceDeclarations(writtenExtensions)));

            static XElement WithoutNamespaceDeclarations(XElement element) {
                var copy = new XElement(element);
                foreach (XAttribute declaration in copy.DescendantsAndSelf().Attributes().Where(attribute => attribute.IsNamespaceDeclaration).ToArray()) {
                    declaration.Remove();
                }
                return copy;
            }
        }
        [Fact]
        public void WorksheetObjectsIgnoreMatchingElementsInUnknownExtensions() {
            using var source = new MemoryStream();
            using (var document = ExcelDocument.Create()) {
                document.AddSheet("Data").Cell(1, 1).SetValue(1);
                document.Save(source);
            }

            XNamespace sheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace spark = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
            XNamespace formula = "http://schemas.microsoft.com/office/excel/2006/main";
            const string partName = "xl/worksheets/sheet1.xml";
            const string shadowUri = "urn:officeimo:test:worksheet-shadow";
            using (var archive = new ZipArchive(source, ZipArchiveMode.Update, leaveOpen: true)) {
                ZipArchiveEntry part = archive.GetEntry(partName)!;
                XDocument worksheet;
                using (Stream input = part.Open()) worksheet = XDocument.Load(input);
                worksheet.Root!.Add(new XElement(sheet + "extLst",
                    new XElement(sheet + "ext", new XAttribute("uri", shadowUri),
                        new XElement(sheet + "pageSetup", new XAttribute("orientation", "landscape")),
                        new XElement(sheet + "headerFooter", new XElement(sheet + "oddHeader", "&LHidden")),
                        new XElement(spark + "sparklineGroups", new XElement(spark + "sparklineGroup",
                            new XElement(spark + "sparklines", new XElement(spark + "sparkline",
                                new XElement(formula + "f", "Data!A1"),
                                new XElement(formula + "sqref", "D1"))))))));
                part.Delete();
                using Stream output = archive.CreateEntry(partName).Open();
                worksheet.Save(output);
            }

            source.Position = 0;
            using var edited = new MemoryStream();
            using (var document = ExcelDocument.Load(source)) {
                ExcelSheet data = document["Data"];
                Assert.Null(data.PageSetup.Orientation);
                Assert.Empty(data.PageSetup.Header.Left.Text);
                Assert.Empty(data.GetSparklines());

                data.PageSetup.Orientation = OfficePageOrientation.Portrait;
                data.PageSetup.Header.Left.Text = "Visible";
                data.AddSparklines("A1", "D1", ExcelSparklineType.Line);
                Assert.Empty(document.OpenXmlSdkFallbacks);
                document.Save(edited);
            }

            edited.Position = 0;
            using var saved = new ZipArchive(edited, ZipArchiveMode.Read);
            using Stream savedPart = saved.GetEntry(partName)!.Open();
            XElement root = XDocument.Load(savedPart).Root!;
            Assert.Equal("portrait", (string?)root.Element(sheet + "pageSetup")?.Attribute("orientation"));
            Assert.Contains("Visible", (string?)root.Element(sheet + "headerFooter")?.Element(sheet + "oddHeader"));
            XElement last = root.Elements().Last();
            Assert.Equal(sheet + "extLst", last.Name);
            XElement shadow = Assert.Single(last.Elements(), element => (string?)element.Attribute("uri") == shadowUri);
            Assert.Equal("&LHidden", (string?)shadow.Element(sheet + "headerFooter")?.Element(sheet + "oddHeader"));
            Assert.Equal("D1", (string?)Assert.Single(shadow.Descendants(formula + "sqref")));
            Assert.Equal(2, last.Elements(sheet + "ext").Count());
        }

        [Fact]
        public void WorksheetViewReadsPendingChangesBeforeDirectSave() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            using var values = new DataTable();
            values.Columns.Add("Name", typeof(string));
            values.Columns.Add("Amount", typeof(int));
            for (int row = 0; row < 100; row++) values.Rows.Add("Item " + row, row);
            sheet.WriteData(values);

            sheet.View.Freeze(2, 1);
            sheet.View.ShowGridlines = false;
            ExcelWorksheetViewInfo pending = sheet.View.Inspect();
            Assert.Equal(2, pending.FrozenRowCount);
            Assert.Equal(1, pending.FrozenColumnCount);
            Assert.False(pending.ShowGridlines);

            using var output = new MemoryStream();
            document.Save(output);
            Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            output.Position = 0;
            using var reopened = ExcelDocument.Load(output);
            ExcelWorksheetViewInfo persisted = reopened["Data"].View.Inspect();
            Assert.Equal(pending.FrozenRowCount, persisted.FrozenRowCount);
            Assert.Equal(pending.FrozenColumnCount, persisted.FrozenColumnCount);
            Assert.False(persisted.ShowGridlines);
        }

        [Fact]
        public void WorksheetViewEditAndPendingCellSurviveFailedSave() {
            using var source = new MemoryStream();
            using (var document = ExcelDocument.Create()) {
                document.AddSheet("Data").Cell(1, 1).SetValue("Original");
                document.Save(source);
            }

            source.Position = 0;
            using var reopened = ExcelDocument.Load(source);
            ExcelSheet sheet = reopened["Data"];
            sheet.Cell(2, 1).SetValue("Pending");
            sheet.View.Freeze(2, 1);
            using var failed = new ThrowAfterBytesWriteStream(128);
            Assert.Throws<IOException>(() => reopened.Save(failed));

            using var retry = new MemoryStream();
            reopened.Save(retry);
            Assert.Empty(reopened.OpenXmlSdkFallbacks);

            retry.Position = 0;
            using var saved = ExcelDocument.Load(retry);
            ExcelSheet savedSheet = saved["Data"];
            Assert.Equal(2, savedSheet.View.FrozenRows);
            Assert.Equal(1, savedSheet.View.FrozenColumns);
            Assert.Equal("Pending", savedSheet.Cell(2, 1).GetValue<string>());
        }
    }
}
