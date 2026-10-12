using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelTabularStyleTests {
        private static readonly DateTime When = new DateTime(2026, 10, 8, 6, 0, 0);
        private static readonly string[] Headers = { "Name", "Amount", "When", "Blank" };

        private static ExcelTabularWriteOptions Options(bool sharedStrings = false, bool references = true) => new() {
            RequireStreaming = true, UseSharedStrings = sharedStrings, IncludeCellReferences = references,
            DateSystem = ExcelDateSystem.NineteenFour,
            Styles = new Dictionary<string, ExcelStyleDefinition> {
                ["Emphasis"] = new() { Bold = true, FontName = "Arial", FontSize = 14,
                    FontColor = OfficeColor.Red, BackgroundColor = OfficeColor.Yellow, WrapText = true,
                    HorizontalAlignment = ExcelHorizontalAlignment.Center, VerticalAlignment = ExcelVerticalAlignment.Top },
                ["Amount"] = new() { Italic = true, NumberFormat = "0.00" },
                ["Plain"] = new()
            },
            DefaultRowStyle = "Emphasis", ColumnStyles = new Dictionary<int, string> { [2] = "Amount" }
        };

        private static ExcelTabularWriteOptions SingleColumnOptions(bool shared = false) {
            var options = Options(shared);
            options.ColumnStyles = null;
            return options;
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public async Task DeclaredDefaultsAndExplicitCellsKeepWholeStylePrecedenceInSyncAndAsync(bool shared, bool references) {
            var options = Options(shared, references);
            using var sync = new MemoryStream();
            using var asyncOutput = new MemoryStream();
            void Write(ExcelDocument.ExcelTabularRowWriter row, int value) {
                if (value == 2) row.SetRowStyle(null);
                row.Write("β & <text>");
                row.Write(12.5);
                if (value == 2) row.SetNextCellStyle("Amount");
                row.Write(When);
                if (value == 1) row.SetNextCellStyle("Plain").WriteBlank();
                else row.Write((string?)null);
            }
            ExcelDocument.WriteRows(sync, new[] { 1, 2 }, Headers, Write, options);
            async IAsyncEnumerable<int> Rows() { yield return 1; await Task.CompletedTask; yield return 2; }
            await ExcelDocument.WriteRowsAsync(asyncOutput, Rows(), Headers, Write, options);

            using var package = SpreadsheetDocument.Open(sync, false);
            using var asyncPackage = SpreadsheetDocument.Open(asyncOutput, false);
            var styles = package.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
            Assert.Equal(styles.OuterXml, asyncPackage.WorkbookPart!.WorkbookStylesPart!.Stylesheet.OuterXml);
            var worksheet = package.WorkbookPart.WorksheetParts.Single().Worksheet;
            Assert.Equal(worksheet.GetFirstChild<SheetData>()!.OuterXml,
                asyncPackage.WorkbookPart.WorksheetParts.Single().Worksheet.GetFirstChild<SheetData>()!.OuterXml);
            Assert.Empty(new OpenXmlValidator().Validate(package));
            Row[] rows = worksheet.Descendants<Row>().ToArray();
            Assert.Null(rows[0].StyleIndex);
            Assert.True(rows[1].CustomFormat!.Value);
            Assert.True(Font(styles, rows[1].StyleIndex!.Value).Bold != null);
            Assert.Null(rows[2].StyleIndex);
            Column column = Assert.Single(worksheet.Descendants<Column>());
            Assert.Equal(2U, column.Min!.Value);
            Assert.Equal(2U, column.Max!.Value);
            Assert.Equal(8.43, column.Width!.Value);
            Assert.Null(column.CustomWidth);
            Assert.NotNull(Font(styles, column.Style!.Value).Italic);
            Cell[] first = rows[1].Elements<Cell>().ToArray();
            Assert.Equal(rows[1].StyleIndex!.Value, first[0].StyleIndex!.Value);
            Assert.Equal(rows[1].StyleIndex!.Value, first[1].StyleIndex!.Value);
            Assert.NotNull(Font(styles, first[2].StyleIndex!.Value).Bold);
            Assert.Equal(164U, Format(styles, first[2].StyleIndex!.Value).NumberFormatId!.Value);
            Assert.Null(first[3].DataType);
            Assert.Null(first[3].CellValue);
            Assert.Null(Font(styles, first[3].StyleIndex!.Value).Bold);
            Cell[] second = rows[2].Elements<Cell>().ToArray();
            Assert.Equal(column.Style!.Value, second[1].StyleIndex!.Value);
            Assert.Equal(column.Style!.Value, second[2].StyleIndex!.Value);
            Assert.Equal(CellValues.String, second[3].DataType!.Value);
            Assert.Equal(shared, package.WorkbookPart.SharedStringTablePart != null);
            using var loaded = ExcelDocument.Load(new MemoryStream(sync.ToArray()));
            Assert.Equal(new[] { "Amount", "Emphasis", "Normal", "Plain" }, loaded.GetNamedStyles().Select(style => style.Name).OrderBy(name => name));
            using var reader = ExcelDocument.OpenDataReader(sync.ToArray());
            Assert.True(reader.Read());
            Assert.Equal("β & <text>", reader.GetString(0));
            Assert.Equal(12.5, reader.GetDouble(1));
            Assert.Equal(When, reader.GetDateTime(2));
            Assert.True(reader.IsDBNull(3));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void TemporalVariantsRetainVisualStyleAndHonorCellValueFormats(bool cellValueFormats) {
            var options = Options();
            options.UseCellValueNumberFormats = cellValueFormats;
            using var output = new MemoryStream();
            ExcelDocument.WriteRows(output, new[] { 1 }, new[] { "Date", "Duration", "Explicit" },
                (row, _) => row.Write(When).Write(TimeSpan.FromHours(30)).SetNextCellStyle("Amount").Write(When), options);
            using var package = SpreadsheetDocument.Open(output, false);
            Stylesheet styles = package.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
            Cell[] cells = package.WorkbookPart.WorksheetParts.Single().Worksheet.Descendants<Row>().Last().Elements<Cell>().ToArray();
            Assert.Equal(cellValueFormats ? 14U : 164U, Format(styles, cells[0].StyleIndex!.Value).NumberFormatId!.Value);
            Assert.Equal(cellValueFormats ? 46U : 165U, Format(styles, cells[1].StyleIndex!.Value).NumberFormatId!.Value);
            Assert.All(cells.Take(2), cell => Assert.NotNull(Font(styles, cell.StyleIndex!.Value).Bold));
            Assert.NotNull(Font(styles, cells[2].StyleIndex!.Value).Italic);
            Assert.Null(Font(styles, cells[2].StyleIndex!.Value).Bold);
            Assert.NotEqual(cellValueFormats ? 14U : 164U, Format(styles, cells[2].StyleIndex!.Value).NumberFormatId!.Value);
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void DataReaderDefaultsApplyToStreamingAndMaterializedOutputIncludingMissingFields(bool shared, bool references) {
            var table = new DataTable();
            table.Columns.Add("Name", typeof(string)); table.Columns.Add("Amount", typeof(double));
            table.Columns.Add("When", typeof(DateTime)); table.Columns.Add("Blank", typeof(string));
            table.Rows.Add("A", 4.5, When, DBNull.Value);
            table.Rows.Add(string.Empty, 0, DBNull.Value, string.Empty);
            var options = Options(shared, references);
            options.RequireStreaming = !shared;
            using var source = table.CreateDataReader();
            using var output = new MemoryStream();
            var result = ExcelDocument.WriteDataReader(output, source, options);
            Assert.Equal(2, result.RowCount);
            Assert.False(source.IsClosed);
            using var package = SpreadsheetDocument.Open(output, false);
            Row[] rows = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Row>().ToArray();
            Assert.All(rows.Skip(1), row => Assert.True(row.CustomFormat!.Value));
            Assert.Null(rows[1].Elements<Cell>().Last().DataType);
            Assert.NotNull(rows[2].Elements<Cell>().Last().DataType);
            Assert.Empty(new OpenXmlValidator().Validate(package));
            using var read = ExcelDocument.OpenDataReader(output.ToArray());
            Assert.True(read.Read()); Assert.Equal(When, read.GetDateTime(2)); Assert.True(read.IsDBNull(3));
            Assert.True(read.Read()); Assert.Equal(string.Empty, read.GetString(0)); Assert.True(read.IsDBNull(2));
            Assert.Equal(string.Empty, read.GetString(3)); Assert.False(read.Read());
        }

        [Fact]
        public void MaterializedDataReaderRetainsCalculatedWidthsTableAndDeclaredDefaults() {
            var table = new DataTable(); table.Columns.Add("Name", typeof(string)); table.Columns.Add("Amount", typeof(double));
            table.Rows.Add(new string('x', 40), 123.5);
            var options = Options(true); options.RequireStreaming = false; options.AutoFit = true; options.CreateTable = true;
            using var source = table.CreateDataReader(); using var output = new MemoryStream();
            ExcelDocument.WriteDataReader(output, source, options);
            using var package = SpreadsheetDocument.Open(output, false);
            WorksheetPart sheet = package.WorkbookPart!.WorksheetParts.Single();
            Column amount = sheet.Worksheet.Descendants<Column>().Single(column => column.Min!.Value == 2);
            Assert.True(amount.Width!.Value > 0); Assert.True(amount.CustomWidth!.Value); Assert.NotNull(amount.Style);
            Assert.Equal("A1:B2", sheet.TableDefinitionParts.Single().Table.Reference!.Value);
            Assert.True(sheet.Worksheet.Descendants<Row>().Last().CustomFormat!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }

        [Fact]
        public void TableRowExportMaterializesSourceOnceAndRetainsRowSelections() {
            int enumerations = 0, callbacks = 0;
            IEnumerable<int> Rows() { Assert.Equal(1, ++enumerations); yield return 1; yield return 2; }
            var options = Options(); options.RequireStreaming = false; options.CreateTable = true;
            using var output = new MemoryStream();
            ExcelDocument.WriteRows(output, Rows(), new[] { "Name", "Amount" }, (row, value) => {
                callbacks++; row.SetRowStyle(value == 1 ? "Emphasis" : null).Write("A").Write(value);
            }, options);
            Assert.Equal(2, callbacks);
            using var package = SpreadsheetDocument.Open(output, false);
            WorksheetPart sheet = package.WorkbookPart!.WorksheetParts.Single();
            Assert.Equal("A1:B3", sheet.TableDefinitionParts.Single().Table.Reference!.Value);
            Row[] rows = sheet.Worksheet.Descendants<Row>().ToArray();
            Assert.NotNull(rows[1].StyleIndex); Assert.Null(rows[2].StyleIndex);
            Assert.NotNull(rows[2].Elements<Cell>().Last().StyleIndex);
        }

        [Fact]
        public async Task InvalidStylesRejectAsyncAndDataReaderBeforeReadingOrPreparingDestination() {
            using var output = new MemoryStream(); output.WriteByte(42);
            bool enumerated = false;
            async IAsyncEnumerable<int> Rows() { enumerated = true; await Task.CompletedTask; yield return 1; }
            var options = SingleColumnOptions(); options.DefaultRowStyle = "Missing";
            await Assert.ThrowsAsync<ArgumentException>(() => ExcelDocument.WriteRowsAsync(output, Rows(), new[] { "A" }, (row, _) => row.Write(1), options));
            Assert.False(enumerated); Assert.Equal(new byte[] { 42 }, output.ToArray()); Assert.Equal(1, output.Position);
            var table = new DataTable(); table.Columns.Add("A", typeof(int)); table.Rows.Add(1);
            using var source = table.CreateDataReader();
            Assert.Throws<ArgumentException>(() => ExcelDocument.WriteDataReader(output, source, options));
            Assert.Equal(new byte[] { 42 }, output.ToArray()); Assert.True(source.Read()); Assert.Equal(1, source.GetInt32(0));
        }

        [Fact]
        public void StandaloneDefinitionsAreUsableNamedStylesAndReplacementPreservesOtherStyles() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 12.5); sheet.CellValue(1, 2, "Other");
            var definition = new ExcelStyleDefinition { Bold = true, BackgroundColor = OfficeColor.Yellow, NumberFormat = "0.00" };
            document.DefineNamedStyle(" Highlight ", definition);
            definition.Bold = false;
            sheet.ApplyNamedStyle("highlight", "A1:A1");
            Assert.True(sheet.GetCellStyle(1, 1).Bold);
            Assert.Equal("0.00", sheet.GetCellStyle(1, 1).NumberFormatCode);
            Assert.False(sheet.GetCellStyle(1, 2).Bold);
            document.DefineNamedStyle("Highlight", new ExcelStyleDefinition { Italic = true });
            Assert.False(sheet.GetCellStyle(1, 1).Bold);
            Assert.True(sheet.GetCellStyle(1, 1).Italic);
            using var loaded = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
            Assert.True(loaded.Sheets[0].GetCellStyle(1, 1).Italic);
        }

        [Theory]
        [InlineData(0)] [InlineData(1)] [InlineData(2)] [InlineData(3)] [InlineData(4)]
        public void InvalidDeclarationsRejectBeforeConsumingSourceOrChangingDestination(int failure) {
            var options = Options();
            switch (failure) {
                case 0: options.DefaultRowStyle = "Missing"; break;
                case 1: options.ColumnStyles = new Dictionary<int, string> { [5] = "Amount" }; break;
                case 2: options.Styles = new Dictionary<string, ExcelStyleDefinition> { ["Normal"] = new() }; break;
                case 3: options.Styles = new Dictionary<string, ExcelStyleDefinition> { ["Emphasis"] = new() { FontSize = double.NaN } }; break;
                case 4: options.Styles = new Dictionary<string, ExcelStyleDefinition> { ["Emphasis"] = new() { NumberFormat = "bad\0format" } }; break;
            }
            using var output = new MemoryStream(); output.Write(new byte[] { 1, 2, 3 }, 0, 3);
            bool enumerated = false;
            IEnumerable<int> Rows() { enumerated = true; yield return 1; }
            Assert.ThrowsAny<ArgumentException>(() => ExcelDocument.WriteRows(output, Rows(), Headers, (row, _) => row.Write(1), options));
            Assert.False(enumerated); Assert.Equal(new byte[] { 1, 2, 3 }, output.ToArray());
            Assert.Equal(3, output.Position);
        }

        [Fact]
        public void SelectionsRejectUnknownAndLateRowStylesWithoutChangingCurrentCell() {
            using var output = new MemoryStream();
            ExcelDocument.WriteRows(output, new[] { 1 }, new[] { "A", "B" }, (row, _) => {
                Assert.Throws<KeyNotFoundException>(() => row.SetRowStyle("Missing"));
                row.Write(1);
                Assert.Throws<InvalidOperationException>(() => row.SetRowStyle("Plain"));
                Assert.Throws<KeyNotFoundException>(() => row.SetNextCellStyle("Missing"));
                row.Write(2);
            }, Options());
            using var package = SpreadsheetDocument.Open(output, false);
            Row row = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Row>().Last();
            Assert.True(row.CustomFormat!.Value); Assert.All(row.Elements<Cell>(), cell => Assert.Equal(row.StyleIndex!.Value, cell.StyleIndex!.Value));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task CancellationDisposesSinglePassSourceAndLeavesDestinationOpen(bool asynchronous) {
            using var output = new MemoryStream();
            using var cancellation = new CancellationTokenSource();
            int enumerations = 0, callbacks = 0; bool disposed = false;
            IEnumerable<int> Rows() { try { Assert.Equal(1, ++enumerations); yield return 1; yield return 2; } finally { disposed = true; } }
            async IAsyncEnumerable<int> AsyncRows() { foreach (int value in Rows()) { await Task.CompletedTask; yield return value; } }
            void Write(ExcelDocument.ExcelTabularRowWriter row, int value) { callbacks++; row.Write(value); cancellation.Cancel(); }
            if (asynchronous) await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ExcelDocument.WriteRowsAsync(output, AsyncRows(), new[] { "Id" }, Write, SingleColumnOptions(), cancellation.Token));
            else Assert.ThrowsAny<OperationCanceledException>(() => ExcelDocument.WriteRows(output, Rows(), new[] { "Id" }, Write, SingleColumnOptions(), cancellation.Token));
            Assert.True(disposed); Assert.Equal(1, callbacks); Assert.True(output.CanWrite);
            output.SetLength(0);
            ExcelDocument.WriteRows(output, new[] { 3 }, new[] { "Id" }, (row, value) => row.Write(value), SingleColumnOptions());
            using var reader = ExcelDocument.OpenDataReader(output.ToArray()); Assert.True(reader.Read()); Assert.Equal(3, reader.GetInt32(0));
        }

        [Fact]
        public void PreCancellationKeepsDestinationAndStyleDefinitionsAreFrozenBeforeEnumeration() {
            using var cancelled = new MemoryStream(); cancelled.WriteByte(42);
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            Assert.ThrowsAny<OperationCanceledException>(() => ExcelDocument.WriteRows(cancelled, new[] { 1 }, new[] { "A" }, (row, _) => row.Write(1), SingleColumnOptions(), cancellation.Token));
            Assert.Equal(new byte[] { 42 }, cancelled.ToArray());
            var options = SingleColumnOptions();
            IEnumerable<int> Rows() { options.Styles!["Emphasis"].Bold = false; yield return 1; }
            using var output = new MemoryStream();
            ExcelDocument.WriteRows(output, Rows(), new[] { "A" }, (row, _) => row.Write(1), options);
            using var package = SpreadsheetDocument.Open(output, false);
            Row row = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Row>().Last();
            Assert.NotNull(Font(package.WorkbookPart.WorkbookStylesPart!.Stylesheet, row.StyleIndex!.Value).Bold);
        }

#if NET8_0_OR_GREATER
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Utf8ValuesUseDeclaredInheritanceAndRemainSinglePass(bool shared) {
            using var output = new MemoryStream(); int enumerations = 0;
            IEnumerable<int> Rows() { Assert.Equal(1, ++enumerations); yield return 1; yield return 2; }
            ExcelDocument.WriteRows(output, Rows(), new[] { "Text" }, (row, _) => row.WriteUtf8("A & β"u8), SingleColumnOptions(shared));
            using var package = SpreadsheetDocument.Open(output, false);
            Assert.All(package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Row>().Skip(1), row => {
                Assert.True(row.CustomFormat!.Value); Assert.Equal(row.StyleIndex!.Value, row.Elements<Cell>().Single().StyleIndex!.Value);
            });
            using var reader = ExcelDocument.OpenDataReader(output.ToArray());
            Assert.True(reader.Read()); Assert.Equal("A & β", reader.GetString(0)); Assert.True(reader.Read()); Assert.Equal("A & β", reader.GetString(0));
        }
#endif
        private static CellFormat Format(Stylesheet styles, uint index) => styles.CellFormats!.Elements<CellFormat>().ElementAt((int)index);
        private static Font Font(Stylesheet styles, uint index) => styles.Fonts!.Elements<Font>().ElementAt((int)Format(styles, index).FontId!.Value);
    }
}
