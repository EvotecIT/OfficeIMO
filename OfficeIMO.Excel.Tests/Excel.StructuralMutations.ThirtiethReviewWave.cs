using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralColumns_NormalizeCellsWithImplicitReferences() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Implicit");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "B");
            Row row = Assert.Single(sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>());
            foreach (Cell cell in row.Elements<Cell>()) cell.CellReference = null;

            ExcelStructuralMutationPlan plan = sheet.PlanInsertColumns(1);

            Assert.Equal(2, plan.AffectedCells);
            plan.Apply();
            Assert.Equal(new[] { "B1", "C1" }, row.Elements<Cell>()
                .Select(cell => cell.CellReference!.Value).ToArray());
            Assert.Equal("A", sheet.CellAt(1, 2).GetValue<string>());
            Assert.Equal("B", sheet.CellAt(1, 3).GetValue<string>());
        }

        [Fact]
        public void Test_CellValueReplacement_ReclaimsOnlyExclusiveInCellImageAssets() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.SetInCellImage(1, 1, TinyPng, altText: "Exclusive");

            sheet.CellValue(1, 1, "Replacement");

            Assert.Empty(document.WorkbookPartRoot.RdRichValueParts);
            Assert.Null(document.WorkbookPartRoot.CellMetadataPart!.Metadata!.GetFirstChild<ValueMetadata>());
            Assert.DoesNotContain(document.WorkbookPartRoot.Parts.Select(pair => pair.OpenXmlPart).OfType<ExtendedPart>(),
                part => part.RelationshipType.EndsWith("/richValueRel", StringComparison.Ordinal));

            sheet.SetInCellImage(2, 1, TinyPng, altText: "Shared");
            sheet.Range("A2").CopyTo("B2");
            sheet.CellValue(2, 1, "Detach one");

            ExcelInCellImage remaining = Assert.Single(sheet.GetInCellImages());
            Assert.Equal("B2", remaining.CellReference);
            Assert.Equal(TinyPng, remaining.Bytes);
            AssertRichImageGraphCounts(document, expected: 1);
        }

        [Fact]
        public void Test_InCellImage_RichValueRelationshipMetadataReadIsBounded() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.SetInCellImage(1, 1, TinyPng);
            ExtendedPart relationshipPart = document.WorkbookPartRoot.Parts
                .Select(pair => pair.OpenXmlPart)
                .OfType<ExtendedPart>()
                .Single(part => part.RelationshipType.EndsWith("/richValueRel", StringComparison.Ordinal));
            using (Stream stream = relationshipPart.GetStream(FileMode.Create, FileAccess.Write)) {
                byte[] block = Enumerable.Repeat((byte)' ', 8192).ToArray();
                for (int index = 0; index <= (16 * 1024 * 1024) / block.Length; index++) {
                    stream.Write(block, 0, block.Length);
                }
            }

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => sheet.GetInCellImages());

            Assert.Contains("Rich-value relationship metadata", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_TableConnectionWithoutQueryPart_TargetsWorksheetParameterRemapping() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelQueryBackedTableInfo source = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "ImportedTableConnection",
                WorksheetName = sheet.Name,
                TableName = "ImportedResults",
                ColumnNames = new[] { "Value" }
            });
            TableDefinitionPart tablePart = Assert.Single(sheet.WorksheetPart.TableDefinitionParts);
            QueryTablePart queryPart = Assert.Single(tablePart.QueryTableParts);
            tablePart.Table!.ConnectionId = source.ConnectionId;
            tablePart.DeletePart(queryPart);
            Assert.Equal(source.ConnectionId, tablePart.Table!.ConnectionId!.Value);
            Connection connection = document.WorkbookPartRoot.ConnectionsPart!.Connections!
                .Elements<Connection>().Single(item => item.Id?.Value == source.ConnectionId);
            var parameter = new Parameter {
                Name = "Input",
                ParameterType = ParameterValues.Cell,
                Cell = "A5"
            };
            connection.Append(new Parameters(parameter) { Count = 1U });

            sheet.InsertColumns(1);

            Assert.Equal("B5", parameter.Cell!.Value);
        }

        [Fact]
        public void Test_RemoveQueryBackedTable_PreservedTableDropsNativeBindingAttributes() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelQueryBackedTableInfo source = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "DetachConnection",
                WorksheetName = sheet.Name,
                TableName = "DetachResults",
                ColumnNames = new[] { "First", "Second" }
            });
            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table!;
            table.ConnectionId = source.ConnectionId;
            Assert.NotNull(table.ConnectionId);
            Assert.All(table.TableColumns!.Elements<TableColumn>(), column => Assert.NotNull(column.QueryTableFieldId));

            Assert.True(document.RemoveQueryBackedTable(source.TableName, preserveTable: true));

            Assert.Equal("A1:B1", sheet.GetTableRange(source.TableName));
            Assert.Null(table.ConnectionId);
            Assert.All(table.TableColumns!.Elements<TableColumn>(), column => Assert.Null(column.QueryTableFieldId));
            Assert.Empty(sheet.WorksheetPart.TableDefinitionParts.Single().QueryTableParts);
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
