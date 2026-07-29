using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text;
using System.Threading;
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tabular.Tests;

public sealed class TabularReaderContracts {
    [Fact]
    public void PublicApi_HasOneReadOnlyEntryPointAndNoBackendReaders() {
        string[] csvTypes = typeof(CsvDocument).Assembly
            .GetExportedTypes()
            .Select(static type => type.FullName ?? type.Name)
            .ToArray();
        string[] excelTypes = typeof(ExcelDocument).Assembly
            .GetExportedTypes()
            .Select(static type => type.FullName ?? type.Name)
            .ToArray();

        Assert.DoesNotContain(csvTypes, static name => name.Contains("CsvDataReader", StringComparison.Ordinal));
        Assert.DoesNotContain(csvTypes, static name => name.Contains("FieldSpanVisitor", StringComparison.Ordinal));
        Assert.DoesNotContain(excelTypes, static name => name.EndsWith(".ExcelDocumentReader", StringComparison.Ordinal));
        Assert.DoesNotContain(excelTypes, static name => name.EndsWith(".ExcelSheetReader", StringComparison.Ordinal));
        Assert.DoesNotContain(excelTypes, static name => name.EndsWith(".ExcelRead", StringComparison.Ordinal));
        Assert.DoesNotContain(excelTypes, static name => name.Contains("ExcelFluentRead", StringComparison.Ordinal));

        MethodInfo[] csvDocumentMethods = typeof(CsvDocument).GetMethods(
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly);
        MethodInfo[] excelDocumentMethods = typeof(ExcelDocument).GetMethods(
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly);
        MethodInfo[] excelSheetMethods = typeof(ExcelSheet).GetMethods(
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly);

        Assert.DoesNotContain(csvDocumentMethods, static method =>
            method.Name is "CreateDataReader" or "ReadRecords" or "ReadRecordsReusable" or
                "ReadRows" or "ReadRowsReusable" or
                "ReadFieldSpans" or "ReadFieldSpansFromText" or
                "ReadRowFieldSpans" or "ReadRowFieldSpansFromText");
        Assert.DoesNotContain(excelDocumentMethods, static method => method.Name == "CreateReader");
        Assert.DoesNotContain(excelSheetMethods, static method =>
            method.Name is "Rows" or "RowsAs" or "RowsAsStream" or "RowsObjects");

        MethodInfo[] canonicalOpenMethods = typeof(TabularReader).GetMethods(
            BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly);
        Assert.NotEmpty(canonicalOpenMethods);
        Assert.All(canonicalOpenMethods, static method => Assert.Equal("Open", method.Name));
    }

    [Fact]
    public void Open_UsesTheSameReaderContractForCsvAndXlsxWithoutRanges() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Tabular.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string csvPath = Path.Combine(directory, "people.csv");
        string xlsxPath = Path.Combine(directory, "people.xlsx");

        try {
            File.WriteAllText(csvPath, "Id,Name\n1,Ada\n2,Grace\n");
            using (var document = ExcelDocument.Create(xlsxPath)) {
                var sheet = document.AddWorksheet("People");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Name");
                sheet.CellValue(2, 1, 1);
                sheet.CellValue(2, 2, "Ada");
                sheet.CellValue(3, 1, 2);
                sheet.CellValue(3, 2, "Grace");
                document.Save();
            }

            AssertRows(csvPath, "people");
            AssertRows(xlsxPath, "People");
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Open_ExposesWorkbookSheetsAsDataReaderResults() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Tabular.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                var first = document.AddWorksheet("First");
                first.CellValue(1, 1, "Value");
                first.CellValue(2, 1, "One");
                var second = document.AddWorksheet("Second");
                second.CellValue(1, 1, "Value");
                second.CellValue(2, 1, "Two");
                document.Save();
            }

            using var reader = TabularReader.Open(path);
            Assert.Equal(new[] { "First", "Second" }, reader.TableNames);
            Assert.Equal("First", reader.TableName);
            Assert.True(reader.Read());
            Assert.Equal("One", reader.GetString(0));
            Assert.True(reader.NextResult());
            Assert.Equal("Second", reader.TableName);
            Assert.True(reader.Read());
            Assert.Equal("Two", reader.GetString(0));
            Assert.False(reader.NextResult());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void Open_XlsbStreamsValuesAndCachedFormulasThroughTheCanonicalReader() {
        string path = GetXlsbFixture("basic-values-formula.xlsb");

        using var reader = TabularReader.Open(path);

        Assert.Equal(TabularFormat.ExcelBinary, reader.Format);
        Assert.Equal("Arkusz1", reader.TableName);
        Assert.Equal("Name", reader.GetName(0));
        Assert.Equal("Amount", reader.GetName(1));
        Assert.True(reader.Read());
        Assert.Equal("Alpha", reader.GetString(0));
        Assert.Equal(42, reader.GetInt32(1));
        Assert.True(reader.Read());
        Assert.Equal(50, reader.GetInt32(1));
        Assert.False(reader.Read());
    }

    [Fact]
    public void Open_XlsbRejectsFormulaCellsWhenCachedResultsAreDisabled() {
        using var reader = TabularReader.Open(
            GetXlsbFixture("basic-values-formula.xlsb"),
            new TabularReadOptions { UseCachedFormulaResult = false });

        Assert.True(reader.Read());
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => reader.Read());
        Assert.Contains("formula-token", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Open_XlsbAppliesDateSystemAndCustomDateStylesWhileStreaming() {
        string path = GetXlsbFixture("styles-dates-formulas.xlsb");

        using var reader = TabularReader.Open(path);

        Assert.Equal("StylesDates", reader.TableName);
        Assert.True(reader.Read());
        Assert.Equal(new DateTime(2024, 2, 29), reader.GetDateTime(0));
        Assert.Equal(1234.6928m, reader.GetDecimal(1));
    }

    [Fact]
    public void Open_CsvStreamsQuotedMultilineEscapedAndLongFields() {
        string longValue = new('x', 70_000);
        string csv = "Id,Description,LongValue\n"
            + "1,\"line one\nline \"\"two\"\"\",\"" + longValue + "\"\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv));

        using var reader = TabularReader.Open(stream, TabularFormat.DelimitedText);

        Assert.True(reader.Read());
        Assert.Equal(1, reader.GetInt32(0));
        Assert.Equal("line one\nline \"two\"", reader.GetString(1));
        Assert.Equal(longValue, reader.GetString(2));
        Assert.False(reader.Read());
    }

    [Fact]
    public void Open_CsvStringColumnsSupportTypedGettersWithoutSchemaInference() {
        Guid identifier = Guid.NewGuid();
        const string csv =
            "Boolean,Byte,Int16,Int32,Int64,Float,Double,Decimal,Date,Guid\n"
            + "true,7,-12,42,9876543210,1.5,2.75,165258.24,2026-07-29,";
        using var stream = new MemoryStream(
            Encoding.UTF8.GetBytes(csv + identifier.ToString("D") + "\n"));

        using var reader = TabularReader.Open(stream, TabularFormat.DelimitedText);

        Assert.True(reader.Read());
        Assert.True(reader.GetBoolean(0));
        Assert.Equal((byte)7, reader.GetByte(1));
        Assert.Equal((short)-12, reader.GetInt16(2));
        Assert.Equal(42, reader.GetInt32(3));
        Assert.Equal(9_876_543_210L, reader.GetInt64(4));
        Assert.Equal(1.5f, reader.GetFloat(5));
        Assert.Equal(2.75d, reader.GetDouble(6));
        Assert.Equal(165258.24m, reader.GetDecimal(7));
        Assert.Equal(new DateTime(2026, 7, 29), reader.GetDateTime(8));
        Assert.Equal(identifier, reader.GetGuid(9));
    }

    [Fact]
    public void Open_CallerOwnedCsvStreamRemainsUsableAfterReaderDisposal() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Id\n1\n"));

        using (var reader = TabularReader.Open(stream, TabularFormat.DelimitedText)) {
            Assert.True(reader.Read());
            Assert.Equal(1, reader.GetInt32(0));
        }

        Assert.True(stream.CanRead);
        stream.Position = 0;
        Assert.Equal((byte)'I', stream.ReadByte());
    }

    [Fact]
    public void Open_SeekableCsvFallbackStartsAtTheCallerPosition() {
        byte[] prefix = Encoding.UTF8.GetBytes("ignored prefix");
        byte[] payload = Encoding.UTF8.GetBytes("Id;Name\n1;Ada\n");
        using var stream = new MemoryStream(prefix.Concat(payload).ToArray());
        stream.Position = prefix.Length;

        using var reader = TabularReader.Open(
            stream,
            TabularFormat.DelimitedText,
            new TabularReadOptions { DetectDelimiter = true });

        Assert.Equal(prefix.Length, stream.Position);
        Assert.Equal("Id", reader.GetName(0));
        Assert.Equal("Name", reader.GetName(1));
        Assert.True(reader.Read());
        Assert.Equal(1, reader.GetInt32(0));
        Assert.Equal("Ada", reader.GetString(1));
        Assert.False(reader.Read());
    }

    [Fact]
    public void Open_CsvObservesCancellationDuringTraversal() {
        using var cancellation = new CancellationTokenSource();
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Id\n1\n2\n"));
        using var reader = TabularReader.Open(
            stream,
            TabularFormat.DelimitedText,
            new TabularReadOptions { CancellationToken = cancellation.Token });

        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => reader.Read());
    }

    [Fact]
    public void Open_PathRejectsInputBeyondTheConfiguredLimit() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Tabular.Limit.{Guid.NewGuid():N}.csv");
        try {
            File.WriteAllText(path, "Id,Name\n1,Ada\n");

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                TabularReader.Open(path, new TabularReadOptions { MaxInputBytes = 4 }));

            Assert.Contains("configured limit", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void Open_SeekableStreamRejectsUnreadInputBeyondTheConfiguredLimit() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Id,Name\n1,Ada\n"));
        stream.Position = 3;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            TabularReader.Open(
                stream,
                TabularFormat.DelimitedText,
                new TabularReadOptions { MaxInputBytes = 4 }));

        Assert.Contains("unread bytes", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(3, stream.Position);
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void Open_XlsbHasRowsRemainsStableAfterTheLastRow() {
        using var reader = TabularReader.Open(GetXlsbFixture("basic-values-formula.xlsb"));

        Assert.True(reader.HasRows);
        while (reader.Read()) {
        }

        Assert.True(reader.HasRows);
    }

    [Fact]
    public void Open_XlsbObservesCancellationDuringTraversal() {
        using var cancellation = new CancellationTokenSource();
        using var reader = TabularReader.Open(
            GetXlsbFixture("basic-values-formula.xlsb"),
            new TabularReadOptions { CancellationToken = cancellation.Token });

        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => reader.Read());
    }

    [Theory]
    [InlineData("A1:Z1000")]
    [InlineData("A1:A2")]
    public void Open_XlsxDiscoversActualBoundsWhenDeclaredDimensionIsStale(string declaredDimension) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Tabular.Dimension.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Name");
                sheet.CellValue(2, 1, 1);
                sheet.CellValue(2, 2, "Ada");
                sheet.CellValue(3, 1, 2);
                sheet.CellValue(3, 2, "Grace");
                document.Save();
            }

            ReplaceWorksheetDimension(path, declaredDimension);

            using var reader = TabularReader.Open(path);
            Assert.Equal(2, reader.FieldCount);
            Assert.Equal("Id", reader.GetName(0));
            Assert.Equal("Name", reader.GetName(1));
            Assert.True(reader.Read());
            Assert.Equal(1, reader.GetInt32(0));
            Assert.Equal("Ada", reader.GetString(1));
            Assert.True(reader.Read());
            Assert.Equal(2, reader.GetInt32(0));
            Assert.Equal("Grace", reader.GetString(1));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void Open_XlsbDiscoversActualColumnsWhenDeclaredDimensionIsStale() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Tabular.Dimension.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            ReplaceXlsbWorksheetLastColumn(path, 0);

            using var reader = TabularReader.Open(path);
            Assert.Equal(2, reader.FieldCount);
            Assert.Equal("Name", reader.GetName(0));
            Assert.Equal("Amount", reader.GetName(1));
            Assert.True(reader.Read());
            Assert.Equal("Alpha", reader.GetString(0));
            Assert.Equal(42, reader.GetInt32(1));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void Open_XlsxUsesConfiguredCultureAndParsesGuidText() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Tabular.Culture.{Guid.NewGuid():N}.xlsx");
        Guid identifier = Guid.NewGuid();
        try {
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Amount");
                sheet.CellValue(1, 2, "Identifier");
                sheet.CellValue(2, 1, "1,5");
                sheet.CellValue(2, 2, identifier.ToString("D"));
                document.Save();
            }

            using var reader = TabularReader.Open(
                path,
                new TabularReadOptions { Culture = CultureInfo.GetCultureInfo("de-DE") });
            Assert.True(reader.Read());
            Assert.Equal(1.5m, reader.GetDecimal(0));
            Assert.Equal(identifier, reader.GetGuid(1));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(TabularFormat.ExcelOpenXml)]
    [InlineData(TabularFormat.ExcelBinary)]
    public void Open_MissingTableDisposesWorkbookOwner(TabularFormat format) {
        string extension = format == TabularFormat.ExcelBinary ? ".xlsb" : ".xlsx";
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Tabular.Missing.{Guid.NewGuid():N}{extension}");
        try {
            if (format == TabularFormat.ExcelBinary) {
                File.Copy(GetXlsbFixture("basic-values-formula.xlsb"), path);
            } else {
                using var document = ExcelDocument.Create(path);
                document.AddWorksheet("Data").CellValue(1, 1, "Value");
                document.Save();
            }

            Assert.Throws<KeyNotFoundException>(() =>
                TabularReader.Open(
                    path,
                    format,
                    new TabularReadOptions { TableName = "Missing" }));

            using FileStream exclusive = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            Assert.True(exclusive.CanWrite);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void Open_XlsxObservesCancellationWhileBufferingInput() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Tabular.Cancel.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                for (int row = 2; row <= 2000; row++) {
                    sheet.CellValue(row, 1, "Value " + row.ToString(CultureInfo.InvariantCulture));
                }
                document.Save();
            }

            using var cancellation = new CancellationTokenSource();
            using var stream = new CancelingReadStream(File.ReadAllBytes(path), cancellation, 1024);
            Assert.Throws<OperationCanceledException>(() =>
                TabularReader.Open(
                    stream,
                    TabularFormat.ExcelOpenXml,
                    new TabularReadOptions { CancellationToken = cancellation.Token }));
            Assert.True(stream.CanRead);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void ReadRecords_UsesTheSameTypedBindingForCsvAndXlsx() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Tabular.Records", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string csvPath = Path.Combine(directory, "orders.csv");
        string xlsxPath = Path.Combine(directory, "orders.xlsx");
        try {
            File.WriteAllText(
                csvPath,
                "Id,Customer,Amount,Placed\n1,Ada,165258.24,2026-07-29\n");
            using (var document = ExcelDocument.Create(xlsxPath)) {
                var sheet = document.AddWorksheet("Orders");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Customer");
                sheet.CellValue(1, 3, "Amount");
                sheet.CellValue(1, 4, "Placed");
                sheet.CellValue(2, 1, 1);
                sheet.CellValue(2, 2, "Ada");
                sheet.CellValue(2, 3, 165258.24m);
                sheet.CellValue(2, 4, new DateTime(2026, 7, 29));
                document.Save();
            }

            AssertRecord(csvPath);
            AssertRecord(xlsxPath);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void ReadRecords_UsesDataMemberNamesForHeadersThatAreNotClrIdentifiers() {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.Tabular.DataMember." + Guid.NewGuid().ToString("N") + ".csv");
        try {
            File.WriteAllText(path, "Order ID,Item Type\n42,Office Supplies\n");

            using var reader = TabularReader.Open(path);
            MappedOrderRecord record = Assert.Single(reader.ReadRecords<MappedOrderRecord>());

            Assert.Equal(42, record.OrderId);
            Assert.Equal("Office Supplies", record.ItemType);
        } finally {
            File.Delete(path);
        }
    }

    private static void AssertRows(string path, string expectedTableName) {
        using var reader = TabularReader.Open(path);
        Assert.Equal(expectedTableName, reader.TableName);
        Assert.Equal(2, reader.FieldCount);
        Assert.Equal("Id", reader.GetName(0));
        Assert.Equal("Name", reader.GetName(1));

        Assert.True(reader.Read());
        Assert.Equal(1, reader.GetInt32(0));
        Assert.Equal("Ada", reader.GetString(1));

        Assert.True(reader.Read());
        Assert.Equal(2, reader.GetInt32(0));
        Assert.Equal("Grace", reader.GetString(1));
        Assert.False(reader.Read());
    }

    private static string GetXlsbFixture(string name) =>
        Path.Combine(
            AppContext.BaseDirectory,
            "Documents",
            "XlsbCorpus",
            "excel-generated",
            name);

    private static void ReplaceWorksheetDimension(string path, string declaredDimension) {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry originalEntry = archive.GetEntry("xl/worksheets/sheet1.xml")
            ?? throw new InvalidDataException("The generated workbook has no first worksheet part.");
        string xml;
        using (var reader = new StreamReader(originalEntry.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true)) {
            xml = reader.ReadToEnd();
        }

        int dimensionStart = xml.IndexOf("<dimension ref=\"", StringComparison.Ordinal);
        Assert.True(dimensionStart >= 0);
        int valueStart = dimensionStart + "<dimension ref=\"".Length;
        int valueEnd = xml.IndexOf('"', valueStart);
        Assert.True(valueEnd > valueStart);
        xml = xml.Substring(0, valueStart) + declaredDimension + xml.Substring(valueEnd);

        originalEntry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(
            "xl/worksheets/sheet1.xml",
            CompressionLevel.Optimal);
        using var writer = new StreamWriter(
            replacement.Open(),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(xml);
    }

    private static void ReplaceXlsbWorksheetLastColumn(string path, uint lastColumn) {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry originalEntry = archive.GetEntry("xl/worksheets/sheet1.bin")
            ?? throw new InvalidDataException("The XLSB fixture has no first worksheet part.");
        byte[] bytes;
        using (Stream input = originalEntry.Open()) {
            using var output = new MemoryStream();
            input.CopyTo(output);
            bytes = output.ToArray();
        }

        bool replaced = false;
        int position = 0;
        while (position < bytes.Length) {
            int firstTypeByte = bytes[position++];
            int type = firstTypeByte & 0x7F;
            if ((firstTypeByte & 0x80) != 0) {
                type |= (bytes[position++] & 0x7F) << 7;
            }

            int size = 0;
            for (int index = 0; index < 4; index++) {
                int current = bytes[position++];
                size |= (current & 0x7F) << (index * 7);
                if ((current & 0x80) == 0 || index == 3) {
                    break;
                }
            }

            if (type == 148) {
                Assert.True(size >= 16);
                WriteUInt32LittleEndian(bytes, position + 12, lastColumn);
                replaced = true;
                break;
            }

            position = checked(position + size);
        }

        Assert.True(replaced);
        originalEntry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(
            "xl/worksheets/sheet1.bin",
            CompressionLevel.Optimal);
        using Stream destination = replacement.Open();
        destination.Write(bytes, 0, bytes.Length);
    }

    private static void WriteUInt32LittleEndian(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }

    private static void AssertRecord(string path) {
        using var reader = TabularReader.Open(path);
        OrderRecord record = Assert.Single(reader.ReadRecords<OrderRecord>());
        Assert.Equal(1, record.Id);
        Assert.Equal("Ada", record.Customer);
        Assert.Equal(165258.24m, record.Amount);
        Assert.Equal(new DateTime(2026, 7, 29), record.Placed);
    }

    private sealed class OrderRecord {
        public OrderRecord() {
        }

        public int Id { get; set; }
        public string Customer { get; set; } = string.Empty;
        public decimal Amount { get; set; }
        public DateTime Placed { get; set; }
    }

    private sealed class CancelingReadStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        private readonly int _maximumReadSize;

        internal CancelingReadStream(
            byte[] bytes,
            CancellationTokenSource cancellation,
            int maximumReadSize)
            : base(bytes, writable: false) {
            _cancellation = cancellation;
            _maximumReadSize = maximumReadSize;
        }

        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, Math.Min(count, _maximumReadSize));
            if (read > 0) {
                _cancellation.Cancel();
            }
            return read;
        }
    }

    [DataContract]
    private sealed class MappedOrderRecord {
        [DataMember(Name = "Order ID")]
        public int OrderId { get; set; }

        [DataMember(Name = "Item Type")]
        public string ItemType { get; set; } = string.Empty;
    }
}
