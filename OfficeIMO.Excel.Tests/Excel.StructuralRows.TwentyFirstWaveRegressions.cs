using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_HonorsEscapedStructuredReferenceBrackets() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Cost [ old");
            sheet.CellAt(2, 1).SetValue(10);
            sheet.CellAt(5, 1).SetValue(20);
            sheet.AddTable(
                "A1:A2",
                hasHeader: true,
                name: "Table1",
                OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            sheet.CellFormula(1, 3, "Table1[Cost '[ old]+Data!A5");

            sheet.InsertRows(5);

            Assert.Equal(
                "Table1[Cost '[ old]+Data!A6",
                sheet.GetFormulaText(1, 3));
        }

        [Fact]
        public void Test_StructuralRows_RemapsQueryTableRefreshSortRanges() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            SortState sortState = AddQueryTableRefreshSortState(
                document,
                sheet,
                "A5:C10");

            sheet.InsertRows(5, 2);

            Assert.Equal("A7:C12", sortState.Reference!.Value);
            Assert.Equal(
                "A7:A12",
                Assert.Single(sortState.Elements<SortCondition>()).Reference!.Value);

            sheet.DeleteRows(7, 2);

            Assert.Equal("A7:C10", sortState.Reference!.Value);
            Assert.Equal(
                "A7:A10",
                Assert.Single(sortState.Elements<SortCondition>()).Reference!.Value);
            var validationErrors = document.ValidateOpenXml();
            Assert.True(
                validationErrors.Count == 0,
                string.Join(System.Environment.NewLine, validationErrors));
        }

        [Fact]
        public void Test_StructuralRows_PreflightsQueryTableRefreshSortRanges() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Keep");
            string reference = $"A{A1.MaxRows}:C{A1.MaxRows}";
            SortState sortState = AddQueryTableRefreshSortState(
                document,
                sheet,
                reference);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.InsertRows(1));

            Assert.Contains("row limit", exception.Message);
            Assert.Equal(reference, sortState.Reference!.Value);
            Assert.Equal("Keep", sheet.CellAt(1, 1).GetValue<string>());
        }

        [Fact]
        public void Test_StructuralRows_AllowsZeroOffsetDrawingBoundaryAtRowLimit() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreateDrawingSheet(document);
            Xdr.TwoCellAnchor anchor = ReplaceWithTwoCellAnchor(
                sheet,
                fromRow: A1.MaxRows - 2,
                toRow: A1.MaxRows - 1,
                toRowOffset: "0",
                Xdr.EditAsValues.TwoCell);

            sheet.InsertRows(A1.MaxRows - 1);

            Assert.Equal(
                (A1.MaxRows - 1).ToString(),
                anchor.FromMarker!.RowId!.Text);
            Assert.Equal(
                A1.MaxRows.ToString(),
                anchor.ToMarker!.RowId!.Text);
        }

        private static SortState AddQueryTableRefreshSortState(
            ExcelDocument document,
            ExcelSheet sheet,
            string reference) {
            var sortState = new SortState(
                new SortCondition {
                    Reference = reference.Replace(":C", ":A")
                }) {
                Reference = reference
            };
            var refresh = new QueryTableRefresh(
                new QueryTableFields { Count = 0U },
                sortState) {
                MinimumVersion = (byte)0,
                NextId = 1U
            };
            ConnectionsPart connectionsPart =
                document.WorkbookPartRoot.ConnectionsPart
                ?? document.WorkbookPartRoot.AddNewPart<ConnectionsPart>();
            connectionsPart.Connections ??= new Connections();
            if (!connectionsPart.Connections
                .Elements<Connection>()
                .Any(connection => connection.Id?.Value == 1U)) {
                connectionsPart.Connections.Append(
                    new Connection {
                        Id = 1U,
                        Name = "Query",
                        Type = 5U,
                        RefreshedVersion = 7
                    });
            }
            connectionsPart.Connections.Save();
            QueryTablePart part = sheet.WorksheetPart.AddNewPart<QueryTablePart>();
            part.QueryTable = new QueryTable(refresh) {
                Name = "Query",
                ConnectionId = 1U
            };
            part.QueryTable.Save();
            return sortState;
        }
    }
}
