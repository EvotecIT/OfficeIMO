using System.ComponentModel;
using System.Globalization;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public sealed class ExcelAutomaticTypedMappingValuesTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AutomaticNativeFieldsPreserveAliasesOrderRoundingAndStructAssignments(bool valueModel) {
        string path = CreateWorkbook(
            "<row r=\"2\"><c r=\"A2\"><v>1.125</v></c><c r=\"B2\" t=\"inlineStr\"><is><t>Żółw 🐢</t></is></c>" +
            "<c r=\"C2\" s=\"1\"><v>45292.25</v></c><c r=\"D2\"><v>7.5</v></c></row>" +
            "<row r=\"3\"><c r=\"A3\"><v>-2.25</v></c><c r=\"B3\" t=\"inlineStr\"><is><t>second</t></is></c>" +
            "<c r=\"C3\" s=\"1\"><v>45293.5</v></c><c r=\"D3\"><v>-1.5</v></c></row>" +
            "<row r=\"4\"><c r=\"A4\"><v>77</v></c><c r=\"C4\" s=\"1\"><v>45294</v></c><c r=\"D4\"><v>5</v></c></row>",
            new[] { "value", "Name", "Date", "Order Id" }, 4);
        try {
            using var reader = ExcelDocument.OpenDataReader(path);
            int count = 0;
            if (valueModel) {
                foreach (NativeValue row in reader.RowsAs<NativeValue>())
                    ValidateNative(row.Name, row.Id, row.Date, row.Value, ++count);
            } else {
                foreach (NativeClass row in reader.RowsAs<NativeClass>())
                    ValidateNative(row.Name, row.Id, row.Date, row.Value, ++count);
            }
            Assert.Equal(3, count);
            Assert.False(reader.IsClosed);
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MixedNativeAndTextDatesKeepRoundTripKindInAutomaticAndExplicitMapping(bool explicitMapping) {
        string path = CreateWorkbook(
            "<row r=\"2\"><c r=\"A2\" s=\"1\"><v>45292.25</v></c></row>" +
            "<row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>2024-01-02T03:04:05.1234567Z</t></is></c></row>",
            new[] { "Date" }, 3);
        try {
            using var reader = ExcelDocument.OpenDataReader(path);
            DateRow[] rows = explicitMapping
                ? reader.RowsAs<DateRow>(map => map.FromColumn<DateTime>("Date", (row, value) => { row.Date = value; return row; })).ToArray()
                : reader.RowsAs<DateRow>().ToArray();
            Assert.Equal(2, rows.Length);
            Assert.Equal(DateTime.FromOADate(45292.25), rows[0].Date);
            DateTime expected = DateTime.Parse("2024-01-02T03:04:05.1234567Z", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
            Assert.Equal(expected, rows[1].Date);
            Assert.Equal(DateTimeKind.Utc, rows[1].Date.Kind);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void SourceTextConversionsAndFormulaTextKeepCanonicalMappingRules() {
        string path = CreateWorkbook(
            "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>1.0</t></is></c>" +
            "<c r=\"B2\" t=\"inlineStr\"><is><t>1</t></is></c><c r=\"C2\"><f>2+3</f><v>5</v></c></row>" +
            "<row r=\"3\"><c r=\"A3\"><v>2</v></c><c r=\"B3\" t=\"b\"><v>0</v></c>" +
            "<c r=\"C3\" t=\"inlineStr\"><is><t>literal</t></is></c></row>",
            new[] { "Id", "Flag", "Formula" }, 3);
        try {
            using (var reader = ExcelDocument.OpenDataReader(path)) {
                TextConversionRow[] rows = reader.RowsAs<TextConversionRow>().ToArray();
                Assert.Equal(2, rows.Length);
                Assert.Equal(1, rows[0].Id);
                Assert.True(rows[0].Flag);
                Assert.Equal("5", rows[0].Formula);
                Assert.Equal(2, rows[1].Id);
                Assert.False(rows[1].Flag);
                Assert.Equal("literal", rows[1].Formula);
            }
            using (var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { UseCachedFormulaResult = false })) {
                TextConversionRow[] rows = reader.RowsAs<TextConversionRow>().ToArray();
                Assert.Equal("2+3", rows[0].Formula);
            }
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("2147483648")]
    [InlineData("not-an-integer")]
    public void InvalidTypedGetterValuesKeepPropertyErrorsAndRedaction(string value) {
        string path = CreateWorkbook($"<row r=\"2\"><c r=\"A2\"><v>{value}</v></c></row>", new[] { "Id" }, 2);
        try {
            using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { MappingErrorValuePolicy = DataMappingErrorValuePolicy.Redact });
            ConstructorProbeRow.Calls = 0;
            DataMappingException error = Assert.Throws<DataMappingException>(() => reader.RowsAs<ConstructorProbeRow>().ToArray());
            Assert.Contains("ConstructorProbeRow.Id", error.Message);
            Assert.DoesNotContain(value, error.Message);
            Assert.Equal(1, ConstructorProbeRow.Calls);
            Assert.False(reader.IsClosed);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void IrregularWidthsMissingAndEmptyFieldsKeepNullableAssignments() {
        string path = CreateWorkbook(
            "<row r=\"2\"><c r=\"B2\"><v>3</v></c></row>" +
            "<row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t></t></is></c>" +
            "<c r=\"B3\"><v>4</v></c><c r=\"C3\" s=\"1\"><v>45292</v></c><c r=\"D3\"><v>1.25</v></c></row>",
            new[] { "Name", "Id", "Date", "Value" }, 3);
        try {
            using var reader = ExcelDocument.OpenDataReader(path);
            NullableRow[] rows = reader.RowsAs<NullableRow>().ToArray();
            Assert.Equal(2, rows.Length);
            Assert.Null(rows[0].Name);
            Assert.Equal(3, rows[0].Id);
            Assert.Null(rows[0].Date);
            Assert.Null(rows[0].Value);
            Assert.Equal(string.Empty, rows[1].Name);
            Assert.Equal(4, rows[1].Id);
            Assert.Equal(DateTime.FromOADate(45292), rows[1].Date);
            Assert.Equal(1.25, rows[1].Value);
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SchemaReplayAndGeneralXmlFallbackKeepTheSameValues(bool inferSchema) {
        string path = CreateWorkbook(
            "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>2024-01-02T03:04:05Z</t></is></c></row>",
            new[] { "Date" }, 2, utf16: !inferSchema);
        try {
            using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { InferSchema = inferSchema, SchemaSampleRows = 1 });
            DateRow row = Assert.Single(reader.RowsAs<DateRow>());
            Assert.Equal(DateTime.Parse("2024-01-02T03:04:05Z", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind), row.Date);
            Assert.Equal(DateTimeKind.Utc, row.Date.Kind);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void DecimalProjectionCultureAndCustomConvertersKeepTheirOriginalInputs() {
        string path = CreateWorkbook(
            "<row r=\"2\"><c r=\"A2\"><v>123456789012345.125</v></c>" +
            "<c r=\"B2\" t=\"inlineStr\"><is><t>1,25</t></is></c></row>", new[] { "Amount", "Value" }, 2);
        try {
            var projection = new ExcelReadOptions { NumericAsDecimal = true, Culture = CultureInfo.GetCultureInfo("pl-PL") };
            decimal expected;
            using (var oracle = ExcelDocument.OpenDataReader(path, projection)) {
                Assert.True(oracle.Read());
                expected = Assert.IsType<decimal>(oracle.GetValue(0));
            }
            using (var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions {
                NumericAsDecimal = true, Culture = CultureInfo.GetCultureInfo("pl-PL")
            })) {
                DecimalRow row = Assert.Single(reader.RowsAs<DecimalRow>());
                Assert.Equal(expected, row.Amount);
                Assert.Equal(1.25, row.Value);
            }
            int typeCalls = 0;
            var converted = new ExcelReadOptions {
                TypeConverter = (raw, type, _) => {
                    typeCalls++;
                    Assert.Equal(7d, Assert.IsType<double>(raw));
                    Assert.Equal(typeof(int), type);
                    return (true, 42);
                },
                CellValueConverter = context => context.RawText == "123456789012345.125"
                    ? new ExcelCellValue(7d) : ExcelCellValue.NotHandled
            };
            using (var reader = ExcelDocument.OpenDataReader(path, converted)) {
                ConverterRow row = Assert.Single(reader.RowsAs<ConverterRow>());
                Assert.Equal(42, row.Id);
            }
            using (var reader = ExcelDocument.OpenDataReader(path, converted)) {
                IdRow row = Assert.Single(reader.RowsAs<IdRow>(map => map.FromColumn<int>("Amount", (record, value) => { record.Id = value; return record; })));
                Assert.Equal(42, row.Id);
            }
            Assert.Equal(2, typeCalls);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void DateStyledNumericProjectionStillUsesWorkbookSerials() {
        string path = CreateWorkbook("<row r=\"2\"><c r=\"A2\" s=\"1\"><v>45292.25</v></c></row>", new[] { "Value" }, 2);
        try {
            using var reader = ExcelDocument.OpenDataReader(path);
            Assert.Equal(45292.25, Assert.Single(reader.RowsAs<NumberRow>()).Value);
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("-0")]
    [InlineData("1.2345678901234567")]
    public void PrimitiveMappingKeepsTheBitsOfTheCurrentPlatformValueConversion(string token) {
        string path = CreateWorkbook($"<row r=\"2\"><c r=\"A2\"><v>{token}</v></c><c r=\"B2\"><v>{token}</v></c></row>",
            new[] { "Double", "Single" }, 2);
        try {
            double expectedDouble;
            float expectedSingle;
            using (var oracle = ExcelDocument.OpenDataReader(path)) {
                Assert.True(oracle.Read());
                expectedDouble = Convert.ToDouble(oracle.GetValue(0), CultureInfo.InvariantCulture);
                expectedSingle = Convert.ToSingle(oracle.GetValue(1), CultureInfo.InvariantCulture);
            }
            using (var reader = ExcelDocument.OpenDataReader(path)) {
                ValidateBits(Assert.Single(reader.RowsAs<PrimitiveBitsRow>()));
            }
            using (var reader = ExcelDocument.OpenDataReader(path)) {
                ValidateBits(Assert.Single(reader.RowsAs<PrimitiveBitsRow>(map => map
                    .FromColumn<double>("Double", (row, value) => { row.Double = value; return row; })
                    .FromColumn<float>("Single", (row, value) => { row.Single = value; return row; }))));
            }

            void ValidateBits(PrimitiveBitsRow row) {
                Assert.Equal(BitConverter.DoubleToInt64Bits(expectedDouble), BitConverter.DoubleToInt64Bits(row.Double));
                Assert.Equal(BitConverter.ToInt32(BitConverter.GetBytes(expectedSingle), 0),
                    BitConverter.ToInt32(BitConverter.GetBytes(row.Single), 0));
            }
        } finally { File.Delete(path); }
    }

    [Fact]
    public void ParallelSchemaReplayKeepsRoundTripKindForSampledAndFollowingTextDates() {
        string[] source = { "2024-01-02T03:04:05.1234567Z", "2024-01-03T04:05:06.7654321Z" };
        string path = CreateWorkbook(
            $"<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>{source[0]}</t></is></c></row>" +
            $"<row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>{source[1]}</t></is></c></row>", new[] { "Date" }, 3);
        try {
            DateTime[] expected = source.Select(text => DateTime.Parse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind)).ToArray();
            foreach (int degree in new[] { 1, 2 }) {
                using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { InferSchema = true, SchemaSampleRows = 1 });
                DateRow[] rows = reader.RowsAsParallel<DateRow>(new ParallelRowMappingOptions { MaxDegreeOfParallelism = degree, BatchSize = 1 }).ToArray();
                Assert.Equal(2, rows.Length);
                for (int index = 0; index < rows.Length; index++) {
                    Assert.Equal(DateTimeKind.Utc, rows[index].Date.Kind);
                    Assert.Equal(expected[index], rows[index].Date);
                }
                Assert.False(reader.IsClosed);
            }
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ParallelCaptureKeepsNumericDateSerialsTextProjectionAndMissingValues(bool inferSchema) {
        string path = CreateWorkbook(
            "<row r=\"2\"><c r=\"A2\" s=\"1\"><v>45292.25</v></c><c r=\"B2\" s=\"1\"><v>45292.25</v></c>" +
            "<c r=\"C2\"><v>1.2345678901234567</v></c><c r=\"D2\" t=\"inlineStr\"><is><t>first</t></is></c></row>" +
            "<row r=\"3\"><c r=\"A3\" s=\"1\"><v>45293.5</v></c><c r=\"B3\" s=\"1\"><v>45293.5</v></c>" +
            "<c r=\"C3\"><v>-1.9876543210987654</v></c></row>", new[] { "Date", "Serial", "NumberText", "Name" }, 3);
        try {
            ParallelCaptureRow[] expected;
            using (var oracle = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { InferSchema = inferSchema, SchemaSampleRows = 1 }))
                expected = oracle.RowsAs<ParallelCaptureRow>().ToArray();
            using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { InferSchema = inferSchema, SchemaSampleRows = 1 });
            ParallelCaptureRow[] rows = reader.RowsAsParallel<ParallelCaptureRow>(new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 }).ToArray();
            Assert.Equal(2, rows.Length);
            for (int index = 0; index < rows.Length; index++) {
                Assert.Equal(index == 0 ? 45292.25 : 45293.5, rows[index].Serial);
                Assert.Equal(expected[index].Date, rows[index].Date);
                Assert.Equal(expected[index].Date.Kind, rows[index].Date.Kind);
                Assert.Equal(expected[index].NumberText, rows[index].NumberText);
                Assert.Equal(index == 0 ? "first" : null, rows[index].Name);
            }
            Assert.False(reader.IsClosed);
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(".xls")]
    [InlineData(".xlsb")]
    public void ParallelLegacyNumericTextKeepsTheCanonicalPlatformFormatting(string extension) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.ParallelCapture.{Guid.NewGuid():N}{extension}");
        try {
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "NumberText");
                sheet.CellValue(2, 1, 1.2345678901234567);
                sheet.CellValue(3, 1, -1.9876543210987654);
                document.Save();
            }
            string[] expected;
            using (var oracle = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { InferSchema = true, SchemaSampleRows = 1 })) {
                var values = new List<string>();
                while (oracle.Read()) values.Add(Convert.ToString(oracle.GetValue(0), CultureInfo.InvariantCulture)!);
                expected = values.ToArray();
            }
            using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { InferSchema = true, SchemaSampleRows = 1 });
            ParallelCaptureRow[] rows = reader.RowsAsParallel<ParallelCaptureRow>(new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 }).ToArray();
            Assert.Equal(expected, rows.Select(row => row.NumberText).ToArray());
            Assert.False(reader.IsClosed);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void InvalidDateSerialKeepsGenericConstructionAndFailureOrder() {
        string path = CreateWorkbook("<row r=\"2\"><c r=\"A2\" s=\"1\"><v>3000000</v></c></row>", new[] { "Date" }, 2);
        try {
            using var reader = ExcelDocument.OpenDataReader(path);
            DateConstructorProbeRow.Calls = 0;
            Assert.Throws<ArgumentException>(() => reader.RowsAs<DateConstructorProbeRow>().ToArray());
            Assert.Equal(1, DateConstructorProbeRow.Calls);
            Assert.False(reader.IsClosed);
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(ExcelDateSystem.NineteenHundred)]
    [InlineData(ExcelDateSystem.NineteenFour)]
    public void MappedDateKeepsTheOriginalSerialForSubsequentGetters(ExcelDateSystem dateSystem) {
        string path = CreateWorkbook("<row r=\"2\"><c r=\"A2\" s=\"1\"><v>1.25</v></c></row>",
            new[] { "Date" }, 2, dateSystem: dateSystem);
        try {
            DateTime expected;
            using (var oracle = ExcelDocument.OpenDataReader(path)) {
                Assert.True(oracle.Read());
                expected = oracle.GetDateTime(0);
            }
            using var reader = ExcelDocument.OpenDataReader(path);
            using IEnumerator<DateRow> rows = reader.RowsAs<DateRow>().GetEnumerator();
            Assert.True(rows.MoveNext());
            Assert.Equal(expected, rows.Current.Date);
            Assert.Equal(1.25, reader.GetDouble(0));
            Assert.Equal(expected, reader.GetDateTime(0));
            Assert.Equal(expected, Assert.IsType<DateTime>(reader.GetValue(0)));
            Assert.Equal(1.25, reader.GetDouble(0));
            Assert.False(rows.MoveNext());
        } finally { File.Delete(path); }
    }

    [Fact]
    public void MalformedDateStyledNumericTokenKeepsTextRoundTripKind() {
        string path = CreateWorkbook("<row r=\"2\"><c r=\"A2\" s=\"1\"><v>2024-01-02T03:04:05Z</v></c></row>",
            new[] { "Date" }, 2);
        try {
            using var reader = ExcelDocument.OpenDataReader(path);
            DateTime date = Assert.Single(reader.RowsAs<DateRow>()).Date;
            Assert.Equal(DateTime.Parse("2024-01-02T03:04:05Z", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind), date);
            Assert.Equal(DateTimeKind.Utc, date.Kind);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void NativeMappingDoesNotRetryOrWrapAPropertySetterFailure() {
        string path = CreateWorkbook("<row r=\"2\"><c r=\"A2\"><v>42</v></c></row>", new[] { "Value" }, 2);
        try {
            using var reader = ExcelDocument.OpenDataReader(path);
            ThrowingSetterRow.Calls = 0;
            Exception original = ThrowingSetterRow.Failure;
            Assert.Same(original, Assert.Throws<FormatException>(() => reader.RowsAs<ThrowingSetterRow>().ToArray()));
            Assert.Equal(1, ThrowingSetterRow.Calls);
            Assert.False(reader.IsClosed);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void MappingCancellationStopsBeforeTheNextNativeRow() {
        string path = CreateWorkbook("<row r=\"2\"><c r=\"A2\"><v>1</v></c></row><row r=\"3\"><c r=\"A3\"><v>2</v></c></row>", new[] { "Id" }, 3);
        try {
            using var cancellation = new CancellationTokenSource();
            using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { CancellationToken = cancellation.Token });
            using IEnumerator<IdRow> rows = reader.RowsAs<IdRow>().GetEnumerator();
            Assert.True(rows.MoveNext());
            Assert.Equal(1, rows.Current.Id);
            cancellation.Cancel();
            OperationCanceledException error = Assert.Throws<OperationCanceledException>(() => rows.MoveNext());
            Assert.Equal(cancellation.Token, error.CancellationToken);
            Assert.False(reader.IsClosed);
        } finally { File.Delete(path); }
    }

#if NET8_0_OR_GREATER
    [Fact]
    public async Task AsyncMappingUsesTheSameNativeAndTextDateContracts() {
        string path = CreateWorkbook(
            "<row r=\"2\"><c r=\"A2\" s=\"1\"><v>45292.25</v></c></row>" +
            "<row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>2024-01-02T03:04:05Z</t></is></c></row>", new[] { "Date" }, 3);
        try {
            using var reader = ExcelDocument.OpenDataReader(path);
            var dates = new List<DateTime>();
            await foreach (DateRow row in reader.RowsAsAsync<DateRow>()) dates.Add(row.Date);
            Assert.Equal(2, dates.Count);
            Assert.Equal(DateTime.FromOADate(45292.25), dates[0]);
            Assert.Equal(DateTimeKind.Utc, dates[1].Kind);
            Assert.Equal(DateTime.Parse("2024-01-02T03:04:05Z", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind), dates[1]);
        } finally { File.Delete(path); }
    }
#endif

    private static void ValidateNative(string? name, int id, DateTime date, double value, int index) {
        Assert.Equal(index == 1 ? "Żółw 🐢" : index == 2 ? "second" : null, name);
        Assert.Equal(index == 1 ? 8 : index == 2 ? -2 : 5, id);
        Assert.Equal(DateTime.FromOADate(index == 1 ? 45292.25 : index == 2 ? 45293.5 : 45294), date);
        Assert.Equal(index == 1 ? 1.125 : index == 2 ? -2.25 : 77, value);
    }

    private static string CreateWorkbook(string rows, string[] headers, int lastRow, bool utf16 = false,
        ExcelDateSystem dateSystem = ExcelDateSystem.NineteenHundred) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.TypedMapping.{Guid.NewGuid():N}.xlsx");
        using (var document = ExcelDocument.Create(path)) {
            document.DateSystem = dateSystem;
            document.AddWorksheet("Data").CellValue(1, 1, "Value");
            document.Save();
        }
        string headerCells = string.Concat(headers.Select((header, index) =>
            $"<c r=\"{(char)('A' + index)}1\" t=\"inlineStr\"><is><t>{header}</t></is></c>"));
        string xml = $"<?xml version=\"1.0\" encoding=\"{(utf16 ? "utf-16" : "utf-8")}\"?>" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            $"<dimension ref=\"A1:{(char)('A' + headers.Length - 1)}{lastRow}\"/><sheetData><row r=\"1\">{headerCells}</row>{rows}</sheetData></worksheet>";
        string styles = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<fonts count=\"1\"><font/></fonts><fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>" +
            "<borders count=\"1\"><border/></borders><cellStyleXfs count=\"1\"><xf/></cellStyleXfs>" +
            "<cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\" applyNumberFormat=\"1\"/></cellXfs></styleSheet>";
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ReplacePart(archive, "xl/worksheets/sheet1.xml", xml, utf16);
        ReplacePart(archive, "xl/styles.xml", styles, false);
        return path;
    }

    private static void ReplacePart(ZipArchive archive, string name, string text, bool utf16) {
        archive.GetEntry(name)?.Delete();
        using Stream output = archive.CreateEntry(name, CompressionLevel.Optimal).Open();
        byte[] bytes = utf16 ? Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(text)).ToArray() : Encoding.UTF8.GetBytes(text);
        output.Write(bytes, 0, bytes.Length);
    }

    public sealed class NativeClass {
        public string? Name { get; set; }
        [DisplayName("Order Id")] public int Id { get; set; }
        public DateTime Date { get; set; }
        public double Value { get; set; }
    }
    public struct NativeValue {
        public string? Name { get; set; }
        [DisplayName("Order Id")] public int Id { get; set; }
        public DateTime Date { get; set; }
        public double Value { get; set; }
    }
    public sealed class DateRow { public DateTime Date { get; set; } }
    public sealed class IdRow { public int Id { get; set; } }
    public sealed class ConverterRow { [DisplayName("Amount")] public int Id { get; set; } }
    public sealed class ConstructorProbeRow {
        internal static int Calls;
        public ConstructorProbeRow() { Calls++; }
        public int Id { get; set; }
    }
    public sealed class DateConstructorProbeRow {
        internal static int Calls;
        public DateConstructorProbeRow() { Calls++; }
        public DateTime Date { get; set; }
    }
    public sealed class NumberRow { public double Value { get; set; } }
    public sealed class PrimitiveBitsRow { public double Double { get; set; } public float Single { get; set; } }
    public sealed class ParallelCaptureRow {
        public DateTime Date { get; set; }
        public double Serial { get; set; }
        public string? NumberText { get; set; }
        public string? Name { get; set; } = "initialized";
    }
    public sealed class DecimalRow { public decimal Amount { get; set; } public double Value { get; set; } }
    public sealed class TextConversionRow { public int Id { get; set; } public bool Flag { get; set; } public string? Formula { get; set; } }
    public sealed class NullableRow {
        public string? Name { get; set; } = "initialized";
        public int? Id { get; set; }
        public DateTime? Date { get; set; }
        public double? Value { get; set; }
    }
    public sealed class ThrowingSetterRow {
        internal static readonly FormatException Failure = new("property setter failure");
        internal static int Calls;
        public int Value { get => 0; set { Calls++; throw Failure; } }
    }
}
