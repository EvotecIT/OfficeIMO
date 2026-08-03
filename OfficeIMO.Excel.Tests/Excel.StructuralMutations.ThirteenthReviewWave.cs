using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ColumnDeletion_VisitsRowsAfterRemovingAnEmptyRow() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Remove");
            sheet.CellValue(2, 2, "Shift");

            sheet.DeleteColumns(1);

            Assert.Equal("Shift", sheet.CellAt(2, 1).GetValue<string>());
            Assert.True(sheet.CellAt(2, 2).GetValue().IsBlank);
        }

        [Theory]
        [InlineData("Data!Other!A1")]
        [InlineData("'Unclosed!A1")]
        [InlineData("Data Name!A1")]
        [InlineData("[Book.xlsx]!A1")]
        public void Test_ReferenceSyntax_RejectsMalformedQualifiers(string text) {
            Assert.False(ExcelReference.TryParse(text, out _));
            Assert.Throws<FormatException>(() => ExcelReference.Parse(text));
        }

        [Fact]
        public void Test_AutoFilterState_AppliesDefaultCellColorMode() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Value");
            sheet.WorksheetPart.Worksheet.Append(new AutoFilter(
                new FilterColumn(new ColorFilter { FormatId = 0U }) { ColumnId = 0U }) {
                Reference = "A1:A2"
            });

            ExcelAutoFilterColumnInfo criterion = Assert.Single(Assert.Single(sheet.GetAutoFilters()).Columns);

            Assert.Equal(ExcelAutoFilterCriteriaKind.Color, criterion.Kind);
            Assert.Equal(true, criterion.CellColor);
        }

        [Fact]
        public void Test_RangeTransfer_BoundsCopiedImagesButMovesWithoutReadingPayloads() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            byte[] imageBytes = new byte[256_000];
            imageBytes[0] = 0x89;
            imageBytes[1] = 0x50;
            imageBytes[2] = 0x4E;
            imageBytes[3] = 0x47;
            sheet.AddImage(1, 1, imageBytes, "image/png", 8, 8);
            var options = new ExcelMutationPlanOptions { MaximumSnapshotCharacters = 128_000 };

            InvalidOperationException copy = Assert.Throws<InvalidOperationException>(() =>
                sheet.CopyRange("A1", "B1", options));

            Assert.Contains("MaximumSnapshotCharacters", copy.Message, StringComparison.Ordinal);
            sheet.MoveRange("A1", "B1", options);
            ExcelImage moved = Assert.Single(sheet.Images);
            Assert.Equal(1, moved.RowIndex);
            Assert.Equal(2, moved.ColumnIndex);
        }

        [Fact]
        public void Test_ColumnPlanning_ChargesConnectionsAndVmlElementsToScanBudget() {
            using (var connectionDocument = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = connectionDocument.AddWorksheet("Data");
                Parameter parameter = AttachCellBackedConnection(connectionDocument, sheet, "A1");
                var options = new ExcelMutationPlanOptions { MaximumScannedElements = 128 };
                Assert.NotNull(sheet.PlanInsertColumns(2, options: options));
                Parameters parameters = (Parameters)parameter.Parent!;
                for (int index = 0; index < 256; index++) {
                    parameters.Append(new Parameter {
                        Name = "Input" + index,
                        ParameterType = ParameterValues.Cell,
                        Cell = "A1"
                    });
                }
                parameters.Count = (uint)parameters.Elements<Parameter>().Count();

                InvalidOperationException connectionBudget = Assert.Throws<InvalidOperationException>(() =>
                    sheet.PlanInsertColumns(2, options: options));

                Assert.Contains("MaximumScannedElements", connectionBudget.Message, StringComparison.Ordinal);
            }

            using (var vmlDocument = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = vmlDocument.AddWorksheet("Data");
                sheet.SetComment(1, 1, "Comment", "Author");
                var options = new ExcelMutationPlanOptions { MaximumScannedElements = 128 };
                Assert.NotNull(sheet.PlanInsertColumns(2, options: options));
                VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
                XDocument markup;
                using (Stream source = vmlPart.GetStream(FileMode.Open, FileAccess.Read)) {
                    markup = XDocument.Load(source);
                }
                XNamespace vml = "urn:schemas-microsoft-com:vml";
                XNamespace excel = "urn:schemas-microsoft-com:office:excel";
                for (int index = 0; index < 128; index++) {
                    markup.Root!.Add(new XElement(vml + "shape",
                        new XElement(excel + "ClientData",
                            new XAttribute("ObjectType", "Note"))));
                }
                using (Stream target = vmlPart.GetStream(FileMode.Create, FileAccess.Write)) {
                    markup.Save(target);
                }

                InvalidOperationException vmlBudget = Assert.Throws<InvalidOperationException>(() =>
                    sheet.PlanInsertColumns(2, options: options));

                Assert.Contains("MaximumScannedElements", vmlBudget.Message, StringComparison.Ordinal);
            }
        }
    }
}
