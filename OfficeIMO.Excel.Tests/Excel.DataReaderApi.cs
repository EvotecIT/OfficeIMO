using System.Data.Common;
using System.Globalization;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Threading;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void DataReaderApi_UsesOpenForSourcesAndCreateForOpenDocuments() {
        Assert.DoesNotContain(
            typeof(ExcelDocument).Assembly.GetExportedTypes(),
            static type => type.Name.EndsWith("DataReader", StringComparison.Ordinal));

        MethodInfo[] methods = typeof(ExcelDocument).GetMethods(
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly);

        Assert.Contains(methods, static method =>
            method.Name == "OpenDataReader"
            && method.IsStatic
            && method.GetParameters().FirstOrDefault()?.ParameterType == typeof(string)
            && method.ReturnType == typeof(DbDataReader));
        Assert.Contains(methods, static method =>
            method.Name == "OpenDataReader"
            && method.IsStatic
            && method.GetParameters().FirstOrDefault()?.ParameterType == typeof(Stream)
            && method.ReturnType == typeof(DbDataReader));
        Assert.Contains(methods, static method =>
            method.Name == "OpenDataReader"
            && method.IsStatic
            && method.GetParameters().FirstOrDefault()?.ParameterType == typeof(byte[])
            && method.ReturnType == typeof(DbDataReader));
        Assert.Contains(methods, static method =>
            method.Name == "CreateDataReader"
            && !method.IsStatic
            && method.ReturnType == typeof(DbDataReader));
    }

    [Fact]
    public void OpenDataReader_ExposesWorksheetsAsOrderedResults() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.DataReader.{Guid.NewGuid():N}.xlsx");
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

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.Read());
            Assert.Equal("One", reader.GetString(0));
            Assert.True(reader.NextResult());
            Assert.True(reader.Read());
            Assert.Equal("Two", reader.GetString(0));
            Assert.False(reader.NextResult());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_SupportsLegacyXlsThroughTheSameEntryPoint() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.DataReader.{Guid.NewGuid():N}.xls");
        try {
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Legacy");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Name");
                sheet.CellValue(2, 1, 7);
                sheet.CellValue(2, 2, "Ada");
                document.Save();
            }

            using (DbDataReader reader = ExcelDocument.OpenDataReader(path)) {
                Assert.Equal("Id", reader.GetName(0));
                Assert.Equal("Name", reader.GetName(1));
                Assert.True(reader.Read());
                Assert.Equal(7, reader.GetInt32(0));
                Assert.Equal("Ada", reader.GetString(1));
                Assert.False(reader.Read());
            }

            byte[] bytes = File.ReadAllBytes(path);
            using var stream = new MemoryStream(bytes, writable: false);
            using DbDataReader streamReader = ExcelDocument.OpenDataReader(stream);
            Assert.True(streamReader.Read());
            Assert.Equal(7, streamReader.GetInt32(0));
            Assert.Equal("Ada", streamReader.GetString(1));
            Assert.True(stream.CanRead);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_LegacyImportUsesTheConfiguredInputLimit() {
        var options = new ExcelReadOptions {
            MaxInputBytes = 128L * 1024L * 1024L
        };

        OfficeIMO.Excel.LegacyXls.LegacyXlsImportOptions importOptions =
            ExcelWorkbookDataReader.CreateLegacyImportOptions(options);

        Assert.Equal(128 * 1024 * 1024, importOptions.MaxInputBytes);
    }

    [Fact]
    public void OpenDataReader_ReadsSeekableWorkbookStreamFromCurrentPositionAndRestoresIt() {
        byte[] workbook = File.ReadAllBytes(GetDataReaderXlsbFixture("basic-values-formula.xlsb"));
        byte[] prefix = Encoding.UTF8.GetBytes("already-consumed-envelope");
        using var stream = new MemoryStream(prefix.Length + workbook.Length);
        stream.Write(prefix, 0, prefix.Length);
        stream.Write(workbook, 0, workbook.Length);
        stream.Position = prefix.Length;

        using DbDataReader reader = ExcelDocument.OpenDataReader(stream);

        Assert.Equal(prefix.Length, stream.Position);
        Assert.True(reader.Read());
        Assert.Equal("Alpha", reader.GetString(0));
        Assert.Equal(42, reader.GetInt32(1));
    }

    [Fact]
    public void OpenDataReader_RejectsUnknownPathExtensionsInsteadOfGuessing() {
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            ExcelDocument.OpenDataReader("workbook.unknown"));

        Assert.Contains(".xlsx", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(".xlsb", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(".xls", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OpenDataReader_SelectsOneWorksheetByName() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.DataReaderSheet.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AddWorksheet("Ignore").CellValue(1, 1, "Ignored");
                var selected = document.AddWorksheet("Data");
                selected.CellValue(1, 1, "Value");
                selected.CellValue(2, 1, "Selected");
                document.Save();
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions { SheetName = "data" });
            Assert.True(reader.Read());
            Assert.Equal("Selected", reader.GetString(0));
            Assert.False(reader.NextResult());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_StreamsXlsbValuesAndCachedFormulas() {
        using DbDataReader reader = ExcelDocument.OpenDataReader(GetDataReaderXlsbFixture("basic-values-formula.xlsb"));

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
    public void OpenDataReader_XlsbRejectsFormulaTokensWhenCachedResultsAreDisabled() {
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            GetDataReaderXlsbFixture("basic-values-formula.xlsb"),
            new ExcelReadOptions { UseCachedFormulaResult = false });

        Assert.True(reader.Read());
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => reader.Read());
        Assert.Contains("formula-token", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OpenDataReader_MissingWorksheetReleasesTheFile() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.MissingSheet.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Value");
                document.Save();
            }

            Assert.Throws<KeyNotFoundException>(() =>
                ExcelDocument.OpenDataReader(path, new ExcelReadOptions { SheetName = "Missing" }));

            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            Assert.True(exclusive.CanWrite);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("A1:Z1000")]
    [InlineData("A1:A2")]
    public void OpenDataReader_XlsxDiscoversActualBoundsWhenDeclaredDimensionIsStale(
        string declaredDimension) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Dimension.{Guid.NewGuid():N}.xlsx");
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

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
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
    public void OpenDataReader_XlsbDiscoversActualColumnsWhenDeclaredDimensionIsStale() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Dimension.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            ReplaceXlsbWorksheetLastColumn(path, 0);

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
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
    public void OpenDataReader_XlsxUsesConfiguredCultureAndParsesGuidText() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Culture.{Guid.NewGuid():N}.xlsx");
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

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions { Culture = CultureInfo.GetCultureInfo("de-DE") });
            Assert.True(reader.Read());
            Assert.Equal(1.5m, reader.GetDecimal(0));
            Assert.Equal(identifier, reader.GetGuid(1));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbMissingWorksheetReleasesTheFile() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.MissingSheet.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            Assert.Throws<KeyNotFoundException>(() =>
                ExcelDocument.OpenDataReader(path, new ExcelReadOptions { SheetName = "Missing" }));

            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            Assert.True(exclusive.CanWrite);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbHasRowsRemainsStableAfterLastRow() {
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            GetDataReaderXlsbFixture("basic-values-formula.xlsb"));

        Assert.True(reader.HasRows);
        while (reader.Read()) {
        }

        Assert.True(reader.HasRows);
    }

    [Fact]
    public void OpenDataReader_XlsxHasRowsRemainsStableAfterLastRow() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.HasRows.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, "One");
                document.Save();
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.HasRows);
            while (reader.Read()) {
            }

            Assert.True(reader.HasRows);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbEnforcesSharedStringLimits() {
        string path = GetDataReaderXlsbFixture("basic-values-formula.xlsb");

        Assert.Throws<InvalidDataException>(() =>
            ExcelDocument.OpenDataReader(path, new ExcelReadOptions { MaxSharedStringItems = 1 }));
        Assert.Throws<InvalidDataException>(() =>
            ExcelDocument.OpenDataReader(path, new ExcelReadOptions { MaxSharedStringItemCharacters = 3 }));
        Assert.Throws<InvalidDataException>(() =>
            ExcelDocument.OpenDataReader(path, new ExcelReadOptions { MaxSharedStringCharacters = 5 }));
    }

    [Fact]
    public void OpenDataReader_PreCancelledLegacyXlsStopsBeforeLoading() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            ExcelDocument.OpenDataReader(
                "not-opened.xls",
                new ExcelReadOptions { CancellationToken = cancellation.Token }));
    }

    [Fact]
    public void OpenDataReader_XlsbObservesCancellationDuringTraversal() {
        using var cancellation = new CancellationTokenSource();
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            GetDataReaderXlsbFixture("basic-values-formula.xlsb"),
            new ExcelReadOptions { CancellationToken = cancellation.Token });

        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => reader.Read());
    }

    [Fact]
    public void OpenDataReader_LegacyXlsObservesCancellationWhileBufferingInput() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Cancel.{Guid.NewGuid():N}.xls");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                for (int row = 2; row <= 2000; row++) {
                    sheet.CellValue(row, 1, "Value " + row.ToString(CultureInfo.InvariantCulture));
                }
                document.Save();
            }

            using var cancellation = new CancellationTokenSource();
            using var stream = new CancelingReadStream(File.ReadAllBytes(path), cancellation, 1024);
            Assert.Throws<OperationCanceledException>(() =>
                ExcelDocument.OpenDataReader(
                    stream,
                    new ExcelReadOptions { CancellationToken = cancellation.Token }));
            Assert.True(stream.CanRead);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsxInfersSchemaIndependentlyOfDataTableInference() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Schema.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, 42);
                document.Save();
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions {
                    InferSchema = true,
                    InferDataTableColumnTypes = false
                });

            Assert.Equal(typeof(double), reader.GetFieldType(0));
            Assert.True(reader.Read());
            Assert.Equal(42, reader.GetInt32(0));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_EmptyHeaderlessXlsxHasNoRows() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Empty.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AddWorksheet("Empty");
                document.Save();
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions { HasHeaderRow = false });

            Assert.Equal(0, reader.FieldCount);
            Assert.False(reader.HasRows);
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbHonorsTheCellValueConverter() {
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            GetDataReaderXlsbFixture("basic-values-formula.xlsb"),
            new ExcelReadOptions {
                CellValueConverter = context =>
                    context.RawText == "42"
                        ? new ExcelCellValue("converted")
                        : ExcelCellValue.NotHandled
            });

        Assert.True(reader.Read());
        Assert.Equal("converted", reader.GetString(1));
        Assert.Equal(typeof(string), reader.GetFieldType(1));
    }

    [Fact]
    public void OpenDataReader_XlsxObservesCancellationWhileBufferingInput() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Cancel.{Guid.NewGuid():N}.xlsx");
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
                ExcelDocument.OpenDataReader(
                    stream,
                    new ExcelReadOptions { CancellationToken = cancellation.Token }));
            Assert.True(stream.CanRead);
        } finally {
            File.Delete(path);
        }
    }

    private static string GetDataReaderXlsbFixture(string name) =>
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
}
