using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        public async Task StreamingRows_InlineTextConsumesSourceOnce(bool asynchronous, bool sharedStrings) {
            string?[] text = ["Alpha", " Alpha & <β>\r ", string.Empty, null, "Alpha", "alpha"];
            DateTime when = new DateTime(2026, 10, 8, 6, 0, 0);
            using var output = new MemoryStream();
            int enumerations = 0;
            int callbacks = 0;
            var options = new ExcelWriterOptions {
                UseSharedStrings = sharedStrings,
                RequireStreaming = true,
                IncludeCellReferences = false,
                DateSystem = ExcelDateSystem.NineteenFour
            };

            IEnumerable<int> Rows() {
                Assert.Equal(1, ++enumerations);
                for (int index = 0; index < text.Length; index++) {
                    Assert.Equal(index, callbacks);
                    yield return index;
                }
            }

            async IAsyncEnumerable<int> AsyncRows() {
                foreach (int index in Rows()) {
                    await Task.CompletedTask;
                    yield return index;
                }
            }

            void WriteRow(ExcelDocument.ExcelTabularRowWriter writer, int index) {
                callbacks++;
                writer.Write(text[index]).Write(index + 1).Write(when);
            }

            ExcelDataSetImportResult result;
            using (var workbookWriter = ExcelDocument.CreateWriter(output, options, leaveOpen: true)) {
                result = asynchronous
                    ? await workbookWriter.WriteTable("Data", AsyncRows(), ["Text", "Id", "When"], WriteRow)
                    : workbookWriter.WriteTable("Data", Rows(), ["Text", "Id", "When"], WriteRow);
            }
            output.Position = 0;
            Assert.Equal(text.Length, result.RowCount);
            Assert.Equal(1, enumerations);
            Assert.Equal(text.Length, callbacks);
            Assert.Equal(sharedStrings, options.UseSharedStrings);

            using (var package = SpreadsheetDocument.Open(output, false)) {
                SharedStringTable? table = package.WorkbookPart!.SharedStringTablePart?.SharedStringTable;
                Assert.Equal(sharedStrings, table != null);
                if (sharedStrings) {
                    Assert.Equal(8U, table!.Count!.Value);
                    Assert.Equal(7U, table.UniqueCount!.Value);
                    Assert.Equal(new[] { "Text", "Id", "When", "Alpha", " Alpha & <β>\r ", string.Empty, "alpha" },
                        table.Elements<SharedStringItem>().Select(item => item.InnerText));
                }
                Cell[] cells = package.WorkbookPart.WorksheetParts.Single().Worksheet.Descendants<Cell>().ToArray();
                Assert.All(cells.Skip(3), cell => Assert.Null(cell.CellReference));
                Assert.All(cells.Where((_, index) => index >= 3 && index % 3 == 2), cell => Assert.NotNull(cell.StyleIndex));
                Assert.True(package.WorkbookPart.Workbook.WorkbookProperties!.Date1904!.Value);
            }

            output.Position = 0;
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(output);
            Assert.Equal(new[] { "Text", "Id", "When" }, Enumerable.Range(0, reader.FieldCount).Select(reader.GetName));
            for (int index = 0; index < text.Length; index++) {
                Assert.True(reader.Read());
                Assert.Equal(text[index] ?? string.Empty, reader.IsDBNull(0) ? string.Empty : reader.GetString(0));
                Assert.Equal(index + 1, reader.GetInt32(1));
                Assert.Equal(when, reader.GetDateTime(2));
            }
            Assert.False(reader.Read());
            Assert.False(reader.NextResult());
        }

        [Fact]
        public void WriteRows_OmittedOptionsKeepInlineTextAndExplicitReferences() {
            using var output = new MemoryStream();
            var destination1 = output;

            var writeOptions1 = new ExcelWriterOptions {
                UseSharedStrings = false,
                CreateTable = false
            };

            using (var workbookWriter1 = ExcelDocument.CreateWriter(destination1, writeOptions1, leaveOpen: true)) {
                workbookWriter1.WriteTable(writeOptions1.SheetName, new[] { "Alpha", "Alpha" }, ["Text"], static (writer, value) => writer.Write(value));
            }

            if (destination1.CanSeek)
                destination1.Position = 0;
            using var package = SpreadsheetDocument.Open(output, false);
            Assert.Null(package.WorkbookPart!.SharedStringTablePart);
            Cell[] cells = package.WorkbookPart.WorksheetParts.Single().Worksheet.Descendants<Cell>().ToArray();
            Assert.Equal(new[] { "A1", "A2", "A3" }, cells.Select(cell => cell.CellReference!.Value));
            Assert.All(cells, cell => Assert.Equal(CellValues.InlineString, cell.DataType!.Value));
        }

        [Fact]
        public void WriteRows_SharedStringsRejectOversizedText() {
            using var output = new MemoryStream();
            Assert.Throws<ArgumentException>(() => {
                var destination2 = output;
                var writeOptions2 = new ExcelWriterOptions {
                    UseSharedStrings = false
                };
                using (var workbookWriter2 = ExcelDocument.CreateWriter(destination2, writeOptions2, leaveOpen: true)) {
                    return workbookWriter2.WriteTable(writeOptions2.SheetName, new[] { new string('x', 32_768) }, ["Text"], static (writer, value) => writer.Write((object)value));
                }
            });
        }
    }
}
