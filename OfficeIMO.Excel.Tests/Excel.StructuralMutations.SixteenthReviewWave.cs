using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ColumnCellShiftAndMove_RemapCellWatchesAndSmartTags() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(2, 2, "Move");
            sheet.CellValue(3, 3, "Keep");

            var movedWatch = new CellWatch { CellReference = "B2" };
            var survivingWatch = new CellWatch { CellReference = "C3" };
            var watches = new CellWatches(movedWatch, survivingWatch);
            sheet.WorksheetPart.Worksheet.Append(watches);

            const string spreadsheetNamespace =
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var tags = new OpenXmlUnknownElement(string.Empty, "cellSmartTags", spreadsheetNamespace);
            tags.SetAttribute(new OpenXmlAttribute(string.Empty, "count", string.Empty, "2"));
            OpenXmlElement movedTag = CreateCellSmartTag(spreadsheetNamespace, "B2");
            OpenXmlElement survivingTag = CreateCellSmartTag(spreadsheetNamespace, "C3");
            tags.Append(movedTag, survivingTag);
            sheet.WorksheetPart.Worksheet.Append(tags);

            sheet.InsertColumns(2);
            Assert.Equal("C2", movedWatch.CellReference!.Value);
            Assert.Equal("C2", GetCellSmartTagReference(movedTag));
            Assert.Equal("D3", survivingWatch.CellReference!.Value);

            sheet.InsertCells("C2", ExcelCellShiftDirection.Right);
            Assert.Equal("D2", movedWatch.CellReference!.Value);
            Assert.Equal("D2", GetCellSmartTagReference(movedTag));

            sheet.MoveRange("D2", "E4");
            Assert.Equal("E4", movedWatch.CellReference!.Value);
            Assert.Equal("E4", GetCellSmartTagReference(movedTag));

            sheet.DeleteCells("E4", ExcelCellShiftDirection.Left);
            Assert.Equal("D3", Assert.Single(watches.Elements<CellWatch>()).CellReference!.Value);
            Assert.Same(survivingTag, Assert.Single(tags.ChildElements));
            Assert.Equal("1", tags.GetAttributes().Single(attribute => attribute.LocalName == "count").Value);
        }

        [Fact]
        public void Test_FileBackedEdit_ReportsPackageSecuritySizeRuleForActiveLimit() {
            string path = Path.Combine(_directoryWithFiles, "FileBackedPackageSecurityLimit.xlsx");
            using (var created = ExcelDocument.Create()) {
                created.AddWorksheet("Data").CellValue(1, 1, "value");
                created.Save(path);
            }

            long sourceBytes = new FileInfo(path).Length;
            var security = new OfficePackageSecurityOptions { MaxPackageBytes = sourceBytes - 1L };
            OfficePackageSecurityException exception = Assert.Throws<OfficePackageSecurityException>(() =>
                ExcelDocument.OpenFileBacked(path, new ExcelLoadOptions {
                    MaxInputBytes = sourceBytes + 1L,
                    PackageSecurity = security
                }));

            Assert.Equal(OfficePackageSecurityRule.PackageSize, exception.Rule);
            Assert.Equal((double) sourceBytes, exception.ObservedValue);
            Assert.Equal((double) (sourceBytes - 1L), exception.Limit);
        }

        [Fact]
        public void Test_FormulaDependencies_ResolveDynamicArraySpillAnchor() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Spill");
            sheet.CellValue(1, 2, 1);
            sheet.CellFormula(1, 1, "SUM(B1#)");

            ExcelFormulaCellInfo formula = Assert.Single(sheet.InspectFormulas().Formulas);

            Assert.Equal(new[] { "Spill!B1" }, formula.Dependencies);
            Assert.DoesNotContain(formula.DependencyIssues, issue =>
                issue.IndexOf("Cannot resolve dependency", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static OpenXmlElement CreateCellSmartTag(string spreadsheetNamespace, string reference) {
            var tag = new OpenXmlUnknownElement(string.Empty, "cellSmartTag", spreadsheetNamespace);
            tag.SetAttribute(new OpenXmlAttribute(string.Empty, "r", string.Empty, reference));
            return tag;
        }

        private static string GetCellSmartTagReference(OpenXmlElement tag) =>
            tag.GetAttributes().Single(attribute => attribute.LocalName == "r").Value;
    }
}
