using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Package;
using OfficeIMO.Excel.Xlsb.Read;
using System.Data.Common;
using System.IO.Compression;
using System.Threading;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void XlsbTabularReader_DisposesItsRecordReaderWhenConstructionFails() {
        using var worksheetPart = new TrackingReadStream(new byte[] { 0x80 });

        Assert.Throws<EndOfStreamException>(() =>
            new XlsbTabularDataReader(
                worksheetPart,
                Array.Empty<string>(),
                Array.Empty<bool>(),
                uses1904DateSystem: false,
                hasHeaderRow: true,
                new ExcelReadOptions(),
                new XlsbImportOptions(),
                new XlsbRecordReadBudget(100),
                new XlsbCellReadBudget(100),
                new XlsbLogicalRowReadBudget(100),
                CancellationToken.None));

        Assert.True(worksheetPart.WasDisposed);
    }

    [Fact]
    public void XlsbTabularReader_EmitsBlankRowsForPhysicalRowGaps() {
        using var worksheetPart = CreateTabularWorksheet(
            (0, 0U),
            (2, 1U));
        using var reader = CreateTabularReader(
            worksheetPart,
            new[] { "Value", "Data" },
            hasHeaderRow: true,
            new XlsbCellReadBudget(10));

        Assert.True(reader.Read());
        Assert.True(reader.IsDBNull(0));
        Assert.True(reader.Read());
        Assert.Equal("Data", reader.GetString(0));
        Assert.False(reader.Read());
    }

    [Fact]
    public void XlsbTabularReader_ChargesSparseRowsAgainstLogicalRowBudget() {
        using var worksheetPart = CreateTabularWorksheet(
            (0, 0U),
            (100, 1U));
        using var reader = CreateTabularReader(
            worksheetPart,
            new[] { "Value", "Data" },
            hasHeaderRow: true,
            new XlsbCellReadBudget(10),
            logicalRowBudget: new XlsbLogicalRowReadBudget(4));

        Assert.True(reader.Read());
        Assert.True(reader.Read());
        Assert.True(reader.Read());
        Assert.Throws<InvalidDataException>(() => reader.Read());
    }

    [Fact]
    public void XlsbTabularReader_SkipsFormattingOnlyRowsBeforeHeaderData() {
        using var worksheetPart = new MemoryStream();
        XlsbRecordWriter.Write(worksheetPart, 129);
        XlsbRecordWriter.Write(
            worksheetPart,
            148,
            CreateWorksheetDimensionPayload(0, 1, 0, 0));
        XlsbRecordWriter.Write(worksheetPart, 145);
        XlsbRecordWriter.Write(
            worksheetPart,
            0,
            CreateTabularRowHeaderPayload(0));
        XlsbRecordWriter.Write(
            worksheetPart,
            0,
            CreateTabularRowHeaderPayload(1));
        XlsbRecordWriter.Write(
            worksheetPart,
            7,
            CreateSharedStringCellPayload(0, 0U));
        XlsbRecordWriter.Write(worksheetPart, 146);
        XlsbRecordWriter.Write(worksheetPart, 130);
        worksheetPart.Position = 0;

        using var reader = CreateTabularReader(
            worksheetPart,
            new[] { "Actual Header" },
            hasHeaderRow: true,
            new XlsbCellReadBudget(10));

        Assert.Equal("Actual Header", reader.GetName(0));
        Assert.False(reader.Read());
    }

    [Fact]
    public void XlsbTabularReader_StopsBeforeFormattingOnlyRowsAfterLastPopulatedRow() {
        using var worksheetPart = new MemoryStream();
        XlsbRecordWriter.Write(worksheetPart, 129);
        XlsbRecordWriter.Write(
            worksheetPart,
            148,
            CreateWorksheetDimensionPayload(0, 100, 0, 0));
        XlsbRecordWriter.Write(worksheetPart, 145);
        XlsbRecordWriter.Write(
            worksheetPart,
            0,
            CreateTabularRowHeaderPayload(0));
        XlsbRecordWriter.Write(
            worksheetPart,
            7,
            CreateSharedStringCellPayload(0, 0U));
        XlsbRecordWriter.Write(
            worksheetPart,
            0,
            CreateTabularRowHeaderPayload(100));
        XlsbRecordWriter.Write(worksheetPart, 146);
        XlsbRecordWriter.Write(worksheetPart, 130);
        worksheetPart.Position = 0;

        using var reader = CreateTabularReader(
            worksheetPart,
            new[] { "Only Value" },
            hasHeaderRow: false,
            new XlsbCellReadBudget(10));

        Assert.True(reader.HasRows);
        Assert.True(reader.Read());
        Assert.Equal("Only Value", reader.GetString(0));
        Assert.False(reader.Read());
    }

    [Fact]
    public void XlsbTabularReader_ChargesWorksheetRecordsOnceAcrossDiscoveryAndDelivery() {
        using var worksheetPart = CreateTabularWorksheet((0, 0U));
        using var reader = CreateTabularReader(
            worksheetPart,
            new[] { "Value" },
            hasHeaderRow: false,
            new XlsbCellReadBudget(10),
            recordBudget: new XlsbRecordReadBudget(7));

        Assert.True(reader.Read());
        Assert.Equal("Value", reader.GetString(0));
        Assert.False(reader.Read());
    }

    [Fact]
    public void XlsbTabularReader_SharesCellBudgetAcrossWorksheetReaders() {
        var cellBudget = new XlsbCellReadBudget(1);
        using (var firstWorksheet = CreateTabularWorksheet((0, 0U)))
        using (var firstReader = CreateTabularReader(
            firstWorksheet,
            new[] { "First" },
            hasHeaderRow: false,
            cellBudget)) {
            Assert.True(firstReader.Read());
            Assert.Equal("First", firstReader.GetString(0));
        }

        using var secondWorksheet = CreateTabularWorksheet((0, 0U));
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            CreateTabularReader(
                secondWorksheet,
                new[] { "Second" },
                hasHeaderRow: false,
                cellBudget));
        Assert.Contains("workbook", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("populated cells", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void XlsbTabularReader_SharesDiscoveryRecordBudgetAcrossWorksheetReaders() {
        var recordBudget = new XlsbRecordReadBudget(9);
        using (var firstWorksheet = CreateTabularWorksheet((0, 0U)))
        using (var firstReader = CreateTabularReader(
            firstWorksheet,
            new[] { "First" },
            hasHeaderRow: false,
            new XlsbCellReadBudget(10),
            recordBudget: recordBudget)) {
            Assert.True(firstReader.HasRows);
        }

        using var secondWorksheet = CreateTabularWorksheet((0, 0U));
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            CreateTabularReader(
                secondWorksheet,
                new[] { "Second" },
                hasHeaderRow: false,
                new XlsbCellReadBudget(10),
                recordBudget: recordBudget));

        Assert.Contains("workbook", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("BIFF12 records", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void XlsbTabularReader_InfersStableSchemaAndReplaysSampledRows() {
        using var worksheetPart = CreateNumericTabularWorksheet(
            (0, 1.25),
            (1, 2.5));
        using var reader = CreateTabularReader(
            worksheetPart,
            Array.Empty<string>(),
            hasHeaderRow: false,
            new XlsbCellReadBudget(10),
            new ExcelReadOptions {
                InferSchema = true,
                SchemaSampleRows = 2
            });

        Assert.Equal(typeof(double), reader.GetFieldType(0));
        Assert.True(reader.Read());
        Assert.Equal(1.25, reader.GetDouble(0));
        Assert.Equal(typeof(double), reader.GetFieldType(0));
        Assert.True(reader.Read());
        Assert.Equal(2.5, reader.GetDouble(0));
        Assert.False(reader.Read());
        Assert.Equal(typeof(double), reader.GetFieldType(0));
    }

    [Fact]
    public void XlsbTabularReader_ReadsEveryValueFromStableDenseRows() {
        using var worksheetPart = CreateTabularWorksheet(
            (0, 0U),
            (1, 1U),
            (2, 2U),
            (3, 3U),
            (4, 4U));
        string[] values = ["Zero", "One", "Two", "Three", "Four"];
        using var reader = CreateTabularReader(
            worksheetPart,
            values,
            hasHeaderRow: false,
            new XlsbCellReadBudget(10));

        var actual = new List<string>();
        while (reader.Read()) {
            actual.Add(reader.GetString(0));
        }

        Assert.Equal(values, actual);
    }

    [Fact]
    public void XlsbTabularReader_FallsBackWhenDenseRowLayoutChanges() {
        using var worksheetPart = CreateLayoutChangingTabularWorksheet();
        using var reader = CreateTabularReader(
            worksheetPart,
            new[] { "Zero", "One", "Two", "Three" },
            hasHeaderRow: false,
            new XlsbCellReadBudget(10));

        object[] expected = ["Zero", "One", "Two", 4.5D, "Three"];
        var actual = new List<object>();
        while (reader.Read()) {
            actual.Add(reader.GetValue(0));
        }

        Assert.Equal(expected, actual);
    }

    [Fact]
    public void XlsbTabularReader_DoesNotExposeTrailingEmptyRowFromDensePlan() {
        using var worksheetPart = CreateTabularWorksheetWithTrailingEmptyRow();
        string[] values = ["Zero", "One", "Two"];
        using var reader = CreateTabularReader(
            worksheetPart,
            values,
            hasHeaderRow: false,
            new XlsbCellReadBudget(10));

        var actual = new List<string>();
        while (reader.Read()) {
            actual.Add(reader.GetString(0));
        }

        Assert.Equal(values, actual);
    }

    [Fact]
    public void XlsbTabularReader_DateCellsRetainNumericTypedAccess() {
        using var worksheetPart = CreateNumericTabularWorksheet((0, 2D));
        using var reader = CreateTabularReader(
            worksheetPart,
            Array.Empty<string>(),
            hasHeaderRow: false,
            new XlsbCellReadBudget(10),
            new ExcelReadOptions {
                TreatDatesUsingNumberFormat = true
            },
            dateStyle: true);

        Assert.True(reader.Read());
        Assert.Equal(2, reader.GetByte(0));
        Assert.Equal(2, reader.GetInt16(0));
        Assert.Equal(2, reader.GetInt32(0));
        Assert.Equal(2L, reader.GetInt64(0));
        Assert.Equal(2F, reader.GetFloat(0));
        Assert.Equal(2D, reader.GetDouble(0));
        Assert.Equal(2M, reader.GetDecimal(0));
        Assert.IsType<DateTime>(reader.GetValue(0));
    }

    [Fact]
    public void XlsbTabularReader_WithoutInferenceKeepsObjectSchemaAcrossMixedRows() {
        using var worksheetPart = CreateMixedTabularWorksheet();
        using var reader = CreateTabularReader(
            worksheetPart,
            new[] { "Text" },
            hasHeaderRow: false,
            new XlsbCellReadBudget(10));

        Assert.Equal(typeof(object), reader.GetFieldType(0));
        Assert.True(reader.Read());
        Assert.Equal(1.25, Assert.IsType<double>(reader.GetValue(0)));
        Assert.Equal(typeof(object), reader.GetFieldType(0));
        Assert.True(reader.Read());
        Assert.Equal("Text", Assert.IsType<string>(reader.GetValue(0)));
        Assert.Equal(typeof(object), reader.GetFieldType(0));
    }

    [Fact]
    public void XlsbTabularReader_NumericAsDecimalFallsBackToDoubleOutsideDecimalRange() {
        using var worksheetPart = CreateNumericTabularWorksheet((0, 1E100));
        using var reader = CreateTabularReader(
            worksheetPart,
            Array.Empty<string>(),
            hasHeaderRow: false,
            new XlsbCellReadBudget(10),
            new ExcelReadOptions { NumericAsDecimal = true });

        Assert.True(reader.Read());
        double value = Assert.IsType<double>(reader.GetValue(0));
        Assert.Equal(1E100, value);
        Assert.Equal(typeof(object), reader.GetFieldType(0));
        Assert.Throws<InvalidCastException>(() => reader.GetDecimal(0));
    }

    [Fact]
    public void XlsbTabularWorkbook_UsesConfiguredAggregateCellLimit() {
        string path = GetDataReaderXlsbFixture("basic-values-formula.xlsb");

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => {
            using DbDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions {
                    HasHeaderRow = false,
                    MaxXlsbCells = 1
                });
            while (reader.Read()) {
            }
        });

        Assert.Contains("1 populated cells", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void XlsbRecordSliceReaders_RejectUnterminatedFourByteSizes() {
        byte[] bytes = { 0x01, 0x80, 0x80, 0x80, 0x80 };
        var arrayReader = new XlsbRecordSliceReader(
            bytes,
            int.MaxValue,
            new XlsbRecordReadBudget(10));
        using var stream = new MemoryStream(bytes, writable: false);
        using var streamReader = new XlsbStreamRecordSliceReader(
            stream,
            int.MaxValue,
            new XlsbRecordReadBudget(10));

        Assert.Throws<InvalidDataException>(() => arrayReader.TryRead(out _));
        Assert.Throws<InvalidDataException>(() => streamReader.TryRead(out _));
    }

    [Fact]
    public void XlsbTabularReader_RejectsWorksheetWithoutOuterBeginSheet() {
        using var worksheetPart = new MemoryStream();
        XlsbRecordWriter.Write(
            worksheetPart,
            148,
            CreateWorksheetDimensionPayload(0, 0, 0, 0));
        XlsbRecordWriter.Write(worksheetPart, 145);
        XlsbRecordWriter.Write(worksheetPart, 146);
        XlsbRecordWriter.Write(worksheetPart, 130);
        worksheetPart.Position = 0;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            CreateTabularReader(
                worksheetPart,
                Array.Empty<string>(),
                hasHeaderRow: false,
                new XlsbCellReadBudget(10)));

        Assert.Contains("BrtBeginSheet", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void XlsbTabularReader_RejectsWorksheetWithoutOuterEndSheet() {
        using var worksheetPart = new MemoryStream();
        XlsbRecordWriter.Write(worksheetPart, 129);
        XlsbRecordWriter.Write(
            worksheetPart,
            148,
            CreateWorksheetDimensionPayload(0, 0, 0, 0));
        XlsbRecordWriter.Write(worksheetPart, 145);
        XlsbRecordWriter.Write(worksheetPart, 146);
        worksheetPart.Position = 0;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            CreateTabularReader(
                worksheetPart,
                Array.Empty<string>(),
                hasHeaderRow: false,
                new XlsbCellReadBudget(10)));

        Assert.Contains("BrtEndSheet", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void XlsbTabularReader_RejectsWorksheetWithoutEndSheetData() {
        using var worksheetPart = new MemoryStream();
        XlsbRecordWriter.Write(worksheetPart, 129);
        XlsbRecordWriter.Write(
            worksheetPart,
            148,
            CreateWorksheetDimensionPayload(0, 0, 0, 0));
        XlsbRecordWriter.Write(worksheetPart, 145);
        XlsbRecordWriter.Write(
            worksheetPart,
            0,
            CreateTabularRowHeaderPayload(0));
        XlsbRecordWriter.Write(
            worksheetPart,
            7,
            CreateSharedStringCellPayload(0, 0U));
        XlsbRecordWriter.Write(worksheetPart, 130);
        worksheetPart.Position = 0;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            CreateTabularReader(
                worksheetPart,
                new[] { "Value" },
                hasHeaderRow: false,
                new XlsbCellReadBudget(10)));

        Assert.Contains("BrtEndSheetData", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void XlsbTabularReader_RejectsWorksheetWithoutBeginSheetData() {
        using var worksheetPart = new MemoryStream();
        XlsbRecordWriter.Write(worksheetPart, 129);
        XlsbRecordWriter.Write(
            worksheetPart,
            148,
            CreateWorksheetDimensionPayload(0, 0, 0, 0));
        XlsbRecordWriter.Write(worksheetPart, 146);
        XlsbRecordWriter.Write(worksheetPart, 130);
        worksheetPart.Position = 0;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            CreateTabularReader(
                worksheetPart,
                Array.Empty<string>(),
                hasHeaderRow: false,
                new XlsbCellReadBudget(10)));

        Assert.Contains("BrtBeginSheetData", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void XlsbTabularReader_RejectsDuplicateBeginSheetData() {
        using var worksheetPart = new MemoryStream();
        XlsbRecordWriter.Write(worksheetPart, 129);
        XlsbRecordWriter.Write(
            worksheetPart,
            148,
            CreateWorksheetDimensionPayload(0, 0, 0, 0));
        XlsbRecordWriter.Write(worksheetPart, 145);
        XlsbRecordWriter.Write(worksheetPart, 145);
        XlsbRecordWriter.Write(worksheetPart, 146);
        XlsbRecordWriter.Write(worksheetPart, 130);
        worksheetPart.Position = 0;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            CreateTabularReader(
                worksheetPart,
                Array.Empty<string>(),
                hasHeaderRow: false,
                new XlsbCellReadBudget(10)));

        Assert.Contains("duplicate", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("BrtBeginSheetData", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void XlsbTabularReader_RejectsUnmatchedEndSheetData() {
        using var worksheetPart = new MemoryStream();
        XlsbRecordWriter.Write(worksheetPart, 129);
        XlsbRecordWriter.Write(
            worksheetPart,
            148,
            CreateWorksheetDimensionPayload(0, 0, 0, 0));
        XlsbRecordWriter.Write(worksheetPart, 145);
        XlsbRecordWriter.Write(worksheetPart, 146);
        XlsbRecordWriter.Write(worksheetPart, 146);
        XlsbRecordWriter.Write(worksheetPart, 130);
        worksheetPart.Position = 0;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            CreateTabularReader(
                worksheetPart,
                Array.Empty<string>(),
                hasHeaderRow: false,
                new XlsbCellReadBudget(10)));

        Assert.Contains("without a matching", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("BrtEndSheetData", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void XlsbTabularReader_RejectsRowThatDoesNotFollowHeader() {
        using var worksheetPart = new MemoryStream();
        XlsbRecordWriter.Write(worksheetPart, 129);
        XlsbRecordWriter.Write(
            worksheetPart,
            148,
            CreateWorksheetDimensionPayload(0, 0, 0, 0));
        XlsbRecordWriter.Write(worksheetPart, 145);
        XlsbRecordWriter.Write(
            worksheetPart,
            0,
            CreateTabularRowHeaderPayload(0));
        XlsbRecordWriter.Write(
            worksheetPart,
            7,
            CreateSharedStringCellPayload(0, 0U));
        XlsbRecordWriter.Write(
            worksheetPart,
            0,
            CreateTabularRowHeaderPayload(0));
        XlsbRecordWriter.Write(
            worksheetPart,
            7,
            CreateSharedStringCellPayload(0, 1U));
        XlsbRecordWriter.Write(worksheetPart, 146);
        XlsbRecordWriter.Write(worksheetPart, 130);
        worksheetPart.Position = 0;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            CreateTabularReader(
                worksheetPart,
                new[] { "Header", "Duplicate" },
                hasHeaderRow: true,
                new XlsbCellReadBudget(10)));

        Assert.Contains("non-increasing row index", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("header row", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void XlsbTabularReader_RejectsCellBeforeRowHeader() {
        using var worksheetPart = new MemoryStream();
        XlsbRecordWriter.Write(worksheetPart, 129);
        XlsbRecordWriter.Write(
            worksheetPart,
            148,
            CreateWorksheetDimensionPayload(0, 0, 0, 0));
        XlsbRecordWriter.Write(worksheetPart, 145);
        XlsbRecordWriter.Write(
            worksheetPart,
            7,
            CreateSharedStringCellPayload(0, 0U));
        XlsbRecordWriter.Write(worksheetPart, 146);
        XlsbRecordWriter.Write(worksheetPart, 130);
        worksheetPart.Position = 0;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            CreateTabularReader(
                worksheetPart,
                new[] { "Value" },
                hasHeaderRow: false,
                new XlsbCellReadBudget(10)));

        Assert.Contains("before a row header", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void XlsbTabularReader_RejectsCellOutsideSheetData() {
        using var worksheetPart = new MemoryStream();
        XlsbRecordWriter.Write(worksheetPart, 129);
        XlsbRecordWriter.Write(
            worksheetPart,
            148,
            CreateWorksheetDimensionPayload(0, 0, 0, 0));
        XlsbRecordWriter.Write(
            worksheetPart,
            7,
            CreateSharedStringCellPayload(0, 0U));
        XlsbRecordWriter.Write(worksheetPart, 145);
        XlsbRecordWriter.Write(worksheetPart, 146);
        XlsbRecordWriter.Write(worksheetPart, 130);
        worksheetPart.Position = 0;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            CreateTabularReader(
                worksheetPart,
                new[] { "Value" },
                hasHeaderRow: false,
                new XlsbCellReadBudget(10)));

        Assert.Contains("outside", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("BrtBeginSheetData", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void XlsbTabularReader_DiscoversHeaderlessColumnsBeyondDeclaredDimension() {
        using var worksheetPart = CreateHeaderlessTabularWorksheet(
            declaredLastColumn: 0,
            (0, 1, 0U));
        using var reader = CreateTabularReader(
            worksheetPart,
            new[] { "Actual" },
            hasHeaderRow: false,
            new XlsbCellReadBudget(10));

        Assert.Equal(2, reader.FieldCount);
        Assert.True(reader.Read());
        Assert.True(reader.IsDBNull(0));
        Assert.Equal("Actual", reader.GetString(1));
    }

    [Theory]
    [InlineData(1, 1, "duplicated")]
    [InlineData(1, 0, "out of order")]
    public void XlsbTabularReader_RejectsDuplicateOrDescendingCellColumns(
        int firstColumn,
        int secondColumn,
        string expectedMessage) {
        using var worksheetPart = CreateHeaderlessTabularWorksheet(
            declaredLastColumn: 1,
            (0, firstColumn, 0U),
            (0, secondColumn, 1U));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            CreateTabularReader(
                worksheetPart,
                new[] { "First", "Second" },
                hasHeaderRow: false,
                new XlsbCellReadBudget(10)));

        Assert.Contains(expectedMessage, exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void XlsbTabularReader_DiscoversHeadedDataColumnsBeyondHeaderAndDeclaredDimension() {
        using var worksheetPart = CreateHeaderlessTabularWorksheet(
            declaredLastColumn: 0,
            (0, 0, 0U),
            (1, 1, 1U));
        using var reader = CreateTabularReader(
            worksheetPart,
            new[] { "Header", "Actual" },
            hasHeaderRow: true,
            new XlsbCellReadBudget(10));

        Assert.Equal(2, reader.FieldCount);
        Assert.Equal("Header", reader.GetName(0));
        Assert.Equal("Column2", reader.GetName(1));
        Assert.True(reader.Read());
        Assert.True(reader.IsDBNull(0));
        Assert.Equal("Actual", reader.GetString(1));
    }

    [Fact]
    public void XlsbPackagePartReader_RejectsPreCancelledPartRead() {
        using var package = new MemoryStream();
        using (var archive = new ZipArchive(package, ZipArchiveMode.Create, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.CreateEntry("xl/sharedStrings.bin");
            using Stream destination = entry.Open();
            destination.Write(new byte[1024], 0, 1024);
        }

        package.Position = 0;
        using var readArchive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var partReader = new XlsbPackagePartReader(readArchive, new XlsbImportOptions());
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            partReader.ReadPart("xl/sharedStrings.bin", cancellation.Token));
    }

    [Fact]
    public void XlsbTabularReader_NullBuffersReportFullFieldLength() {
        using (var worksheetPart = CreateTabularWorksheet((0, 0U)))
        using (var reader = CreateTabularReader(
            worksheetPart,
            new[] { "abcdef" },
            hasHeaderRow: false,
            new XlsbCellReadBudget(10))) {
            Assert.True(reader.Read());
            Assert.Equal(6, reader.GetChars(0, 3, null, 0, 0));
        }

        using (var worksheetPart = CreateTabularWorksheet((0, 0U)))
        using (var reader = CreateTabularReader(
            worksheetPart,
            new[] { "ignored" },
            hasHeaderRow: false,
            new XlsbCellReadBudget(10),
            new ExcelReadOptions {
                CellValueConverter = static _ => new ExcelCellValue(new byte[] { 1, 2, 3, 4 })
            })) {
            Assert.True(reader.Read());
            Assert.Equal(4, reader.GetBytes(0, 2, null, 0, 0));
        }
    }

    [Fact]
    public void XlsbTabularReader_SchemaBudgetCountsRowsActuallyBuffered() {
        using var worksheetPart = CreateWideNumericTabularWorksheet(fieldCount: 1000);
        using var reader = CreateTabularReader(
            worksheetPart,
            Array.Empty<string>(),
            hasHeaderRow: false,
            new XlsbCellReadBudget(10),
            new ExcelReadOptions {
                InferSchema = true,
                SchemaSampleRows = 1024,
                MaxDataReaderBufferedCells = 1000
            });

        Assert.Equal(1000, reader.FieldCount);
        Assert.True(reader.Read());
        Assert.Equal(1.25, reader.GetDouble(0));
        Assert.Equal(2.5, reader.GetDouble(999));
        Assert.False(reader.Read());
    }

    [Fact]
    public void XlsbTabularReader_ConverterDbNullMatchesReaderAndSchemaNullContract() {
        using var worksheetPart = CreateTabularWorksheet(
            (0, 0U),
            (1, 1U));
        using var reader = CreateTabularReader(
            worksheetPart,
            new[] { "missing", "value" },
            hasHeaderRow: false,
            new XlsbCellReadBudget(10),
            new ExcelReadOptions {
                InferSchema = true,
                CellValueConverter = static context =>
                    context.RawText == "0"
                        ? new ExcelCellValue(DBNull.Value)
                        : new ExcelCellValue(42)
            });

        Assert.Equal(typeof(int), reader.GetFieldType(0));
        Assert.True(reader.Read());
        Assert.Equal(DBNull.Value, reader.GetValue(0));
        Assert.True(reader.IsDBNull(0));
        Assert.Throws<InvalidCastException>(() => reader.GetString(0));
        Assert.Throws<InvalidCastException>(() => reader.GetChars(0, 0, null, 0, 0));
        Assert.True(reader.Read());
        Assert.Equal(42, reader.GetInt32(0));
        Assert.False(reader.IsDBNull(0));
    }

    [Fact]
    public void ExcelDocumentReader_DisposesOpenedDocumentWhenInitializationFails() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.ReaderDispose.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Value");
                document.Save();
            }

            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() =>
                ExcelDocumentReader.Open(
                    path,
                    new ExcelReadOptions { CancellationToken = cancellation.Token }));

            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            Assert.True(exclusive.CanRead);
        } finally {
            File.Delete(path);
        }
    }

    private static XlsbTabularDataReader CreateTabularReader(
        Stream worksheetPart,
        IReadOnlyList<string> sharedStrings,
        bool hasHeaderRow,
        XlsbCellReadBudget cellBudget,
        ExcelReadOptions? options = null,
        XlsbRecordReadBudget? recordBudget = null,
        bool dateStyle = false,
        XlsbLogicalRowReadBudget? logicalRowBudget = null) =>
        new(
            worksheetPart,
            sharedStrings,
            new[] { dateStyle },
            uses1904DateSystem: false,
            hasHeaderRow,
            options ?? new ExcelReadOptions(),
            new XlsbImportOptions(),
            recordBudget ?? new XlsbRecordReadBudget(100),
            cellBudget,
            logicalRowBudget ?? new XlsbLogicalRowReadBudget(100),
            CancellationToken.None);

    private static MemoryStream CreateTabularWorksheet(params (int RowIndex, uint SharedStringIndex)[] rows) {
        var stream = new MemoryStream();
        int lastRow = rows.Length == 0 ? 0 : rows.Max(static row => row.RowIndex);
        XlsbRecordWriter.Write(stream, 129);
        XlsbRecordWriter.Write(stream, 148, CreateWorksheetDimensionPayload(0, lastRow, 0, 0));
        XlsbRecordWriter.Write(stream, 145);
        foreach ((int rowIndex, uint sharedStringIndex) in rows) {
            XlsbRecordWriter.Write(stream, 0, CreateTabularRowHeaderPayload(rowIndex));
            XlsbRecordWriter.Write(stream, 7, CreateSharedStringCellPayload(0, sharedStringIndex));
        }

        XlsbRecordWriter.Write(stream, 146);
        XlsbRecordWriter.Write(stream, 130);
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateNumericTabularWorksheet(params (int RowIndex, double Value)[] rows) {
        var stream = new MemoryStream();
        int lastRow = rows.Length == 0 ? 0 : rows.Max(static row => row.RowIndex);
        XlsbRecordWriter.Write(stream, 129);
        XlsbRecordWriter.Write(stream, 148, CreateWorksheetDimensionPayload(0, lastRow, 0, 0));
        XlsbRecordWriter.Write(stream, 145);
        foreach ((int rowIndex, double value) in rows) {
            XlsbRecordWriter.Write(stream, 0, CreateTabularRowHeaderPayload(rowIndex));
            XlsbRecordWriter.Write(stream, 5, CreateRealCellPayload(0, value));
        }

        XlsbRecordWriter.Write(stream, 146);
        XlsbRecordWriter.Write(stream, 130);
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateMixedTabularWorksheet() {
        var stream = new MemoryStream();
        XlsbRecordWriter.Write(stream, 129);
        XlsbRecordWriter.Write(stream, 148, CreateWorksheetDimensionPayload(0, 1, 0, 0));
        XlsbRecordWriter.Write(stream, 145);
        XlsbRecordWriter.Write(stream, 0, CreateTabularRowHeaderPayload(0));
        XlsbRecordWriter.Write(stream, 5, CreateRealCellPayload(0, 1.25));
        XlsbRecordWriter.Write(stream, 0, CreateTabularRowHeaderPayload(1));
        XlsbRecordWriter.Write(stream, 7, CreateSharedStringCellPayload(0, 0U));
        XlsbRecordWriter.Write(stream, 146);
        XlsbRecordWriter.Write(stream, 130);
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateLayoutChangingTabularWorksheet() {
        var stream = new MemoryStream();
        XlsbRecordWriter.Write(stream, 129);
        XlsbRecordWriter.Write(stream, 148, CreateWorksheetDimensionPayload(0, 4, 0, 0));
        XlsbRecordWriter.Write(stream, 145);
        for (int rowIndex = 0; rowIndex < 3; rowIndex++) {
            XlsbRecordWriter.Write(stream, 0, CreateTabularRowHeaderPayload(rowIndex));
            XlsbRecordWriter.Write(stream, 7, CreateSharedStringCellPayload(0, (uint)rowIndex));
        }

        XlsbRecordWriter.Write(stream, 0, CreateTabularRowHeaderPayload(3));
        XlsbRecordWriter.Write(stream, 5, CreateRealCellPayload(0, 4.5D));
        XlsbRecordWriter.Write(stream, 0, CreateTabularRowHeaderPayload(4));
        XlsbRecordWriter.Write(stream, 7, CreateSharedStringCellPayload(0, 3U));
        XlsbRecordWriter.Write(stream, 146);
        XlsbRecordWriter.Write(stream, 130);
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateTabularWorksheetWithTrailingEmptyRow() {
        var stream = new MemoryStream();
        XlsbRecordWriter.Write(stream, 129);
        XlsbRecordWriter.Write(stream, 148, CreateWorksheetDimensionPayload(0, 3, 0, 0));
        XlsbRecordWriter.Write(stream, 145);
        for (int rowIndex = 0; rowIndex < 3; rowIndex++) {
            XlsbRecordWriter.Write(stream, 0, CreateTabularRowHeaderPayload(rowIndex));
            XlsbRecordWriter.Write(stream, 7, CreateSharedStringCellPayload(0, (uint)rowIndex));
        }

        XlsbRecordWriter.Write(stream, 0, CreateTabularRowHeaderPayload(3));
        XlsbRecordWriter.Write(stream, 146);
        XlsbRecordWriter.Write(stream, 130);
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateHeaderlessTabularWorksheet(
        int declaredLastColumn,
        params (int RowIndex, int Column, uint SharedStringIndex)[] cells) {
        var stream = new MemoryStream();
        int lastRow = cells.Length == 0 ? 0 : cells.Max(static cell => cell.RowIndex);
        XlsbRecordWriter.Write(stream, 129);
        XlsbRecordWriter.Write(stream, 148, CreateWorksheetDimensionPayload(0, lastRow, 0, declaredLastColumn));
        XlsbRecordWriter.Write(stream, 145);
        foreach (IGrouping<int, (int RowIndex, int Column, uint SharedStringIndex)> row in
                 cells.GroupBy(static cell => cell.RowIndex).OrderBy(static row => row.Key)) {
            XlsbRecordWriter.Write(stream, 0, CreateTabularRowHeaderPayload(row.Key));
            foreach ((int _, int column, uint sharedStringIndex) in row) {
                XlsbRecordWriter.Write(stream, 7, CreateSharedStringCellPayload(column, sharedStringIndex));
            }
        }

        XlsbRecordWriter.Write(stream, 146);
        XlsbRecordWriter.Write(stream, 130);
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateWideNumericTabularWorksheet(int fieldCount) {
        var stream = new MemoryStream();
        XlsbRecordWriter.Write(stream, 129);
        XlsbRecordWriter.Write(stream, 148, CreateWorksheetDimensionPayload(0, 0, 0, fieldCount - 1));
        XlsbRecordWriter.Write(stream, 145);
        XlsbRecordWriter.Write(stream, 0, CreateTabularRowHeaderPayload(0));
        XlsbRecordWriter.Write(stream, 5, CreateRealCellPayload(0, 1.25));
        XlsbRecordWriter.Write(stream, 5, CreateRealCellPayload(fieldCount - 1, 2.5));
        XlsbRecordWriter.Write(stream, 146);
        XlsbRecordWriter.Write(stream, 130);
        stream.Position = 0;
        return stream;
    }

    private static byte[] CreateWorksheetDimensionPayload(
        int firstRow,
        int lastRow,
        int firstColumn,
        int lastColumn) {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream);
        writer.Write(firstRow);
        writer.Write(lastRow);
        writer.Write(firstColumn);
        writer.Write(lastColumn);
        return stream.ToArray();
    }

    private static byte[] CreateTabularRowHeaderPayload(int rowIndex) {
        byte[] payload = new byte[17 + 16 * 8];
        using var stream = new MemoryStream(payload, writable: true);
        using var writer = new BinaryWriter(stream);
        writer.Write(rowIndex);
        stream.Position = 13;
        writer.Write(16U);
        for (uint segment = 0; segment < 16U; segment++) {
            writer.Write(segment * 1024U);
            writer.Write(segment * 1024U + 1023U);
        }
        return payload;
    }

    private static byte[] CreateSharedStringCellPayload(int column, uint sharedStringIndex) {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream);
        writer.Write(column);
        writer.Write(0U);
        writer.Write(sharedStringIndex);
        return stream.ToArray();
    }

    private static byte[] CreateRealCellPayload(int column, double value) {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream);
        writer.Write(column);
        writer.Write(0U);
        writer.Write(value);
        return stream.ToArray();
    }

    private sealed class TrackingReadStream : MemoryStream {
        internal TrackingReadStream(byte[] bytes) : base(bytes, writable: false) {
        }

        internal bool WasDisposed { get; private set; }

        protected override void Dispose(bool disposing) {
            WasDisposed = true;
            base.Dispose(disposing);
        }
    }
}
