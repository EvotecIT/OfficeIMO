using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_MutationSnapshotRoot_RejectsSerializedXmlAtTheConfiguredBoundary() {
            var worksheet = new Worksheet(new SheetData(new Row(new Cell {
                CellReference = "A1",
                CellValue = new CellValue(new string('x', 32_768)),
                DataType = CellValues.String
            })));

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                ExcelSheet.MeasureMutationSnapshotRoot(
                    worksheet,
                    remainingCharacters: 256,
                    maximumCharacters: 256));

            Assert.Contains("MaximumSnapshotCharacters (256)", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_TableResize_RelocatesTotalsCellsToTheNewBottomRow() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Region");
            sheet.CellValue(1, 2, "Amount");
            sheet.CellValue(2, 1, "East");
            sheet.CellValue(2, 2, 10d);
            sheet.CellValue(3, 1, "West");
            sheet.CellValue(3, 2, 20d);
            sheet.AddTable("A1:B4", true, "Sales", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table!;
            table.TotalsRowShown = true;
            table.TotalsRowCount = 1U;
            table.AutoFilter!.Reference = "A1:B3";
            sheet.CellValue(4, 1, "Total");
            sheet.CellAt(4, 2).SetFormula("SUBTOTAL(109,Sales[Amount])");
            table.Save();

            sheet.ResizeTable("Sales", "A1:B6");

            Assert.True(sheet.CellAt(4, 1).GetValue().IsBlank);
            Assert.Equal("Total", sheet.CellAt(6, 1).GetValue<string>());
            Assert.Equal("SUBTOTAL(109,Sales[Amount])", sheet.GetFormulaText(6, 2));

            sheet.ResizeTable("Sales", "A1:B3");

            Assert.Equal("Total", sheet.CellAt(3, 1).GetValue<string>());
            Assert.Equal("SUBTOTAL(109,Sales[Amount])", sheet.GetFormulaText(3, 2));
            Assert.True(sheet.CellAt(6, 1).GetValue().IsBlank);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public async Task Test_QueryRefresh_PreservesTotalsImageAfterDiscardedImageRenumbersMetadata() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelQueryBackedTableInfo source = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "ImageTotals",
                WorksheetName = sheet.Name,
                TableName = "ImageResults",
                ColumnNames = new[] { "Region", "Badge" }
            });
            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table!;
            table.Reference = "A1:B3";
            table.TotalsRowShown = true;
            table.TotalsRowCount = 1U;
            table.AutoFilter!.Reference = "A1:B2";
            sheet.SetInCellImage(1, 1, TinyPng, altText: "Discarded header");
            sheet.SetInCellImage(3, 2, TinyPng, altText: "Preserved total");
            table.Save();
            var host = new StubQueryHost(new ExcelQueryExecutionResult(
                new[] { "Region", "Badge" },
                new IReadOnlyList<object?>[] {
                    new object?[] { "East", "Ready" },
                    new object?[] { "West", "Ready" }
                }));

            await document.RefreshQueryAsync(
                source.TableName,
                host,
                new ExcelQueryExecutionPolicy { AllowExecution = true });

            ExcelInCellImage image = Assert.Single(sheet.GetInCellImages());
            Assert.Equal("B4", image.CellReference);
            Assert.Equal("Preserved total", image.AltText);
            Assert.Equal(TinyPng, image.Bytes);
            Assert.Equal("Region", sheet.CellAt(1, 1).GetValue<string>());
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
