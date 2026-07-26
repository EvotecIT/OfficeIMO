using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_PreflightsWorkbookConsolidationSources() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            var source = new DataReference {
                Sheet = "Data",
                Reference = $"A{A1.MaxRows}"
            };
            summary.WorksheetPart.Worksheet.Append(
                new DataConsolidate(
                    new DataReferences(source) { Count = 1U }) {
                    Function = DataConsolidateFunctionValues.Sum
                });

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => data.InsertRows(1));

            Assert.Contains("row limit", exception.Message);
            Assert.Equal($"A{A1.MaxRows}", source.Reference!.Value);
        }

        [Fact]
        public void Test_StructuralRows_RemovesEmptyConsolidationSourceContainer() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellAt(2, 1).SetValue(1);
            var consolidate = new DataConsolidate(
                new DataReferences(
                    new DataReference {
                        Sheet = "Data",
                        Reference = "A2"
                    }) {
                    Count = 1U
                }) {
                Function = DataConsolidateFunctionValues.Sum
            };
            summary.WorksheetPart.Worksheet.Append(consolidate);

            data.DeleteRows(2);

            Assert.Null(consolidate.GetFirstChild<DataReferences>());
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_StructuralRows_RemovesScenariosWithoutInputs() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue(1);
            sheet.WorksheetPart.Worksheet.Append(
                new Scenarios(
                    new Scenario(
                        new InputCells {
                            CellReference = "A2",
                            Val = "10"
                        }) {
                        Name = "Only",
                        Count = 1U
                    }));

            sheet.DeleteRows(2);

            Assert.Null(sheet.WorksheetPart.Worksheet.GetFirstChild<Scenarios>());
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_StructuralRows_RemovesEmptyRowBreakContainer() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue(1);
            sheet.AddManualRowPageBreak(2);

            sheet.DeleteRows(2);

            Assert.Null(sheet.WorksheetPart.Worksheet.GetFirstChild<RowBreaks>());
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_StructuralRows_RemovesEmptySparklineExtension() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue(1);
            sheet.CellAt(2, 2).SetValue(2);
            sheet.AddSparklines("A2:B2", "C2");

            sheet.DeleteRows(2);

            Assert.Empty(sheet.WorksheetPart.Worksheet.Descendants<X14.Sparkline>());
            Assert.Empty(sheet.WorksheetPart.Worksheet.Descendants<X14.SparklineGroup>());
            Assert.Empty(sheet.WorksheetPart.Worksheet.Descendants<X14.SparklineGroups>());
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_StructuralRows_RemapsWebPublishedRanges() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(5, 1).SetValue(1);
            WorkbookPart workbookPart = sheet.WorksheetPart.GetParentParts().OfType<WorkbookPart>().Single();
            var item = new WebPublishItem {
                Id = 1U,
                DivId = "Data_1",
                SourceType = WebSourceValues.Range,
                SourceObject = "Data",
                SourceRef = "A5:B6",
                DestinationFile = "published.htm"
            };
            workbookPart.Workbook.Append(new WebPublishItems(item) { Count = 1U });

            sheet.InsertRows(5);

            Assert.Equal("A6:B7", item.SourceRef!.Value);
            Assert.Equal(1U, workbookPart.Workbook.GetFirstChild<WebPublishItems>()!.Count!.Value);
        }

        [Fact]
        public void Test_StructuralRows_RemapsCellSmartTagReferences() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(5, 1).SetValue(1);
            const string spreadsheetNamespace =
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var tags = new OpenXmlUnknownElement(string.Empty, "smartTags", spreadsheetNamespace);
            var tag = new OpenXmlUnknownElement(string.Empty, "cellSmartTag", spreadsheetNamespace);
            tag.SetAttribute(new OpenXmlAttribute(string.Empty, "r", string.Empty, "A5"));
            tags.Append(tag);
            sheet.WorksheetPart.Worksheet.Append(tags);

            sheet.InsertRows(5);

            OpenXmlAttribute reference = tag.GetAttributes()
                .Single(attribute => attribute.LocalName == "r");
            Assert.Equal("A6", reference.Value);
        }
    }
}
