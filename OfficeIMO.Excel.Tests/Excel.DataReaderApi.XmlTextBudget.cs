using System.Data.Common;
using System.Threading;
using Xunit;

namespace OfficeIMO.Excel.Tests {
    public partial class Excel {
        [Fact]
        public void DataReader_XmlTextBudgetRequiresPositiveLimitAndPreservesClonedOptions() {
            var options = new ExcelReadOptions();
            Assert.Equal(32L * 1024L * 1024L, options.MaxXmlDataReaderBufferedCharacters);
            Assert.Throws<ArgumentOutOfRangeException>(() => options.MaxXmlDataReaderBufferedCharacters = 0);
            Assert.Throws<ArgumentOutOfRangeException>(() => options.MaxXmlDataReaderBufferedCharacters = -1);
            options.MaxXmlDataReaderBufferedCharacters = 5;
            Assert.Equal(5L, options.Clone().MaxXmlDataReaderBufferedCharacters);
        }

        [Theory]
        [InlineData("scalar")]
        [InlineData("value")]
        [InlineData("null")]
        [InlineData("values")]
        public void DataReader_XmlTextBudgetRejectsAggregateRowAcrossGetterFamilies(string firstAccess) {
            string path = CreateXmlTextBudgetWorkbook(
                "<row r=\"2\"><c r=\"A2\"><v>123</v></c>"
                + "<c r=\"B2\" t=\"inlineStr\"><is><t>abc</t></is></c></row>", columns: 2);
            try {
                var options = new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = 5 };
                using DbDataReader reader = firstAccess == "values"
                    ? ExcelDocument.OpenDataReader(File.ReadAllBytes(path), options.Clone())
                    : ExcelDocument.OpenDataReader(path, options);
                Assert.True(reader.Read());
                AssertXmlTextBudgetFailure(() => ReadXmlTextBudgetFirstValue(reader, firstAccess));
                AssertXmlTextBudgetReaderIsTerminal(reader);
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData("numeric", "scalar")]
        [InlineData("raw", "value")]
        [InlineData("inline", "null")]
        [InlineData("formula", "values")]
        public void DataReader_XmlTextBudgetRejectsSingleValueLargerThanReadBuffer(string kind, string firstAccess) {
            // Each value is 4,097 characters: rejection must cover chunked XML text reads.
            string text = kind switch {
                "numeric" => "0." + new string('0', 4095),
                "formula" => string.Concat(Enumerable.Repeat("1+", 2048)) + "1",
                _ => new string('x', 4097)
            };
            string cell = kind switch {
                "numeric" => "<c r=\"A2\"><v>" + text + "</v></c>",
                "raw" => "<c r=\"A2\" t=\"str\"><v>" + text + "</v></c>",
                "inline" => "<c r=\"A2\" t=\"inlineStr\"><is><t>" + text + "</t></is></c>",
                _ => "<c r=\"A2\"><f>" + text + "</f><v>1</v></c>"
            };
            string path = CreateXmlTextBudgetWorkbook("<row r=\"2\">" + cell + "</row>");
            try {
                using DbDataReader reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions {
                    MaxXmlDataReaderBufferedCharacters = 4096,
                    UseCachedFormulaResult = kind != "formula"
                });
                Assert.True(reader.Read());
                AssertXmlTextBudgetFailure(() => ReadXmlTextBudgetFirstValue(reader, firstAccess));
                AssertXmlTextBudgetReaderIsTerminal(reader);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void DataReader_XmlTextBudgetAllowsExactBoundaryAndResetsAfterHeaderAndPhysicalRows() {
            string path = CreateXmlTextBudgetWorkbook(
                "<row r=\"2\"><c r=\"A2\"><v>123</v></c>"
                + "<c r=\"B2\" t=\"inlineStr\"><is><t>abc</t></is></c></row>"
                + "<row r=\"3\"><c r=\"A3\"><v>456</v></c>"
                + "<c r=\"B3\" t=\"inlineStr\"><is><t>def</t></is></c></row>", columns: 2,
                headerCells: "<c r=\"A1\" t=\"inlineStr\"><is><t>One</t></is></c>"
                    + "<c r=\"B1\" t=\"inlineStr\"><is><t>Two</t></is></c>");
            try {
                using DbDataReader reader = ExcelDocument.OpenDataReader(path,
                    new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = 6 });
                Assert.Equal("One", reader.GetName(0));
                Assert.Equal("Two", reader.GetName(1));
                Assert.True(reader.Read());
                Assert.Equal(123, reader.GetInt32(0));
                Assert.Equal("abc", reader.GetString(1));
                var values = new object[2];
                for (int repeat = 0; repeat < 2; repeat++) {
                    Assert.Equal(2, reader.GetValues(values));
                    Assert.Equal(new object[] { 123D, "abc" }, values);
                    Assert.False(reader.IsDBNull(1));
                }
                Assert.True(reader.Read());
                Assert.Equal(2, reader.GetValues(values));
                Assert.Equal(new object[] { 456D, "def" }, values);
                Assert.True(reader.Read());
                Assert.True(reader.IsDBNull(0));
                Assert.True(reader.IsDBNull(1));
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData("<c r=\"A2\" t=\"inlineStr\"><is><r><t xml:space=\"preserve\"> A</t></r>"
            + "<r><t><![CDATA[<&]]>B&amp;🐢</t></r><r><t xml:space=\"preserve\"> \t </t></r></is></c>",
            " A<&B&🐢 \t ")]
        [InlineData("<c r=\"A2\" t=\"str\"><v xml:space=\"preserve\"> A<![CDATA[<&]]>B&amp;🐢 \t </v></c>",
            " A<&B&🐢 \t ")]
        [InlineData("<c r=\"A2\" t=\"inlineStr\"><is><t xml:space=\"preserve\"> \t\n </t></is></c>", " \t\n ")]
        public void DataReader_XmlTextBudgetCountsDecodedRichTextCDataAndWhitespace(string cell, string expected) {
            string path = CreateXmlTextBudgetWorkbook("<row r=\"2\">" + cell + "</row>");
            try {
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path,
                    new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = expected.Length })) {
                    Assert.True(reader.Read());
                    Assert.Equal(expected, reader.GetString(0));
                    Assert.Equal(expected, reader.GetValue(0));
                }
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path,
                    new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = expected.Length - 1 })) {
                    Assert.True(reader.Read());
                    AssertXmlTextBudgetFailure(() => reader.GetString(0));
                    AssertXmlTextBudgetReaderIsTerminal(reader);
                }
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData("cachedNumber", 4)]
        [InlineData("cachedNumberValueFirst", 4)]
        [InlineData("cachedBoolean", 4)]
        [InlineData("formula", 3)]
        [InlineData("date", 8)]
        public void DataReader_XmlTextBudgetCountsFormulaCachedAndDateTextBeforeScalarConversion(string kind, long characters) {
            string cell = kind switch {
                "cachedNumber" => "<c r=\"A2\"><f>1+2</f><v>3</v></c>",
                "cachedNumberValueFirst" => "<c r=\"A2\"><v>3</v><f>1+2</f></c>",
                "cachedBoolean" => "<c r=\"A2\" t=\"b\"><f>1=1</f><v>1</v></c>",
                "formula" => "<c r=\"A2\"><f>1+2</f><v>12345</v></c>",
                _ => "<c r=\"A2\" s=\"1\"><v>45351.25</v></c>"
            };
            string path = CreateXmlTextBudgetWorkbook("<row r=\"2\">" + cell + "</row>");
            try {
                if (kind == "date") {
                    ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes(
                        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                        + "<cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/></cellXfs></styleSheet>"));
                }
                var options = new ExcelReadOptions {
                    MaxXmlDataReaderBufferedCharacters = characters,
                    UseCachedFormulaResult = kind != "formula",
                    TreatDatesUsingNumberFormat = true
                };
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path, options)) {
                    Assert.True(reader.Read());
                    if (kind == "formula") {
                        Assert.Equal("1+2", reader.GetString(0));
                    } else if (kind == "cachedBoolean") {
                        Assert.True(reader.GetBoolean(0));
                    } else if (kind == "date") {
                        DateTime expected = ExcelDateSystemConverter.FromSerial(45351.25D, ExcelDateSystem.NineteenHundred);
                        Assert.Equal(expected, reader.GetDateTime(0));
                        Assert.Equal(45351.25D, reader.GetDouble(0));
                        Assert.Equal(expected, reader.GetValue(0));
                    } else {
                        Assert.Equal(3, reader.GetInt32(0));
                    }
                }
                options.MaxXmlDataReaderBufferedCharacters = characters - 1;
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path, options)) {
                    Assert.True(reader.Read());
                    AssertXmlTextBudgetFailure(() => {
                        if (kind == "date") reader.GetDateTime(0);
                        else if (kind == "cachedBoolean") reader.GetBoolean(0);
                        else if (kind == "formula") reader.GetString(0);
                        else reader.GetInt32(0);
                    });
                    AssertXmlTextBudgetReaderIsTerminal(reader);
                }
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void DataReader_XmlTextBudgetCountsSharedStringIndexAndResolvedText() {
            const string expected = "A&B🐢";
            string path = CreateXmlTextBudgetWorkbook(
                "<row r=\"2\"><c r=\"A2\" t=\"s\"><v>0</v></c></row>");
            try {
                ReplaceZipEntry(path, "xl/sharedStrings.xml", Encoding.UTF8.GetBytes(
                    "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">"
                    + "<si><r><t>A&amp;B</t></r><r><t>🐢</t></r></si></sst>"));
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path,
                    new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = expected.Length + 1 })) {
                    Assert.True(reader.Read());
                    Assert.Equal(expected, reader.GetString(0));
                    Assert.Equal(expected, reader.GetValue(0));
                }
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path,
                    new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = expected.Length })) {
                    Assert.True(reader.Read());
                    AssertXmlTextBudgetFailure(() => reader.IsDBNull(0));
                    AssertXmlTextBudgetReaderIsTerminal(reader);
                }
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void DataReader_XmlTextBudgetBoundsLeadingZeroSharedStringIndexesDuringOpen(bool multipleSheets) {
            string index = new string('0', 12);
            string path = CreateXmlTextBudgetWorkbook(
                "<row r=\"2\"><c r=\"A2\" t=\"s\"><v>" + index + "</v></c></row>"
                + "<row r=\"3\"><c r=\"A3\" t=\"s\"><v>" + index + "</v></c></row>",
                multipleSheets: multipleSheets);
            try {
                ReplaceZipEntry(path, "xl/sharedStrings.xml", Encoding.UTF8.GetBytes(
                    "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\" uniqueCount=\"1\">"
                    + "<si><t/></si></sst>"));
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path,
                    new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = index.Length })) {
                    Assert.True(reader.Read());
                    Assert.Equal(string.Empty, reader.GetString(0));
                    Assert.True(reader.Read());
                    Assert.Equal(string.Empty, reader.GetString(0));
                }
                AssertXmlTextBudgetFailure(() => {
                    using DbDataReader reader = ExcelDocument.OpenDataReader(path,
                        new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = index.Length - 1 });
                });
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void DataReader_XmlTextBudgetBoundsSharedFormulaFollowerWhitespaceDuringOpen() {
            string whitespace = new string(' ', 12);
            string path = CreateXmlTextBudgetWorkbook(
                "<row r=\"2\"><c r=\"A2\"><f t=\"shared\" si=\"0\" ref=\"A2:A4\">1+2</f><v>3</v></c></row>"
                + "<row r=\"3\"><c r=\"A3\"><f t=\"shared\" si=\"0\" xml:space=\"preserve\">"
                + whitespace + "</f><v>7</v></c></row>"
                + "<row r=\"4\"><c r=\"A4\"><f t=\"shared\" si=\"0\" xml:space=\"preserve\">"
                + whitespace + "</f><v>8</v></c></row>");
            try {
                // Preparation reads formula text to recognize a follower, without materializing numeric v text.
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path,
                    new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = whitespace.Length })) {
                    Assert.Equal("A", reader.GetName(0));
                    Assert.True(reader.Read());
                    Assert.Equal(3, reader.GetInt32(0));
                    Assert.True(reader.Read());
                }
                // Materializing the follower also counts its one-character cached value.
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path,
                    new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = whitespace.Length + 1 })) {
                    Assert.True(reader.Read());
                    Assert.Equal(3, reader.GetInt32(0));
                    Assert.True(reader.Read());
                    Assert.Equal(7, reader.GetInt32(0));
                    Assert.True(reader.Read());
                    Assert.Equal(8, reader.GetInt32(0));
                }
                AssertXmlTextBudgetFailure(() => {
                    using DbDataReader reader = ExcelDocument.OpenDataReader(path,
                        new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = whitespace.Length - 1 });
                });
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData("expanded", 11)]
        [InlineData("sameReference", 3)]
        [InlineData("equalDistinctString", 6)]
        public void DataReader_XmlTextBudgetCountsDistinctConverterTextAndRetainsRawReferences(string kind, long characters) {
            string path = CreateXmlTextBudgetWorkbook(
                "<row r=\"2\"><c r=\"A2\" t=\"str\"><v>abc</v></c></row>");
            try {
                var options = new ExcelReadOptions {
                    MaxXmlDataReaderBufferedCharacters = characters,
                    CellValueConverter = context => context.RawText != "abc" ? ExcelCellValue.NotHandled
                        : new ExcelCellValue(kind == "expanded" ? "expanded"
                            : kind == "sameReference" ? context.RawText : new string(context.RawText.ToCharArray()))
                };
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path, options)) {
                    Assert.True(reader.Read());
                    Assert.Equal(kind == "expanded" ? "expanded" : "abc", reader.GetString(0));
                    Assert.False(reader.IsDBNull(0));
                }
                options.MaxXmlDataReaderBufferedCharacters = characters - 1;
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path, options)) {
                    Assert.True(reader.Read());
                    AssertXmlTextBudgetFailure(() => reader.GetValue(0));
                    AssertXmlTextBudgetReaderIsTerminal(reader);
                }
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(false, 3)]
        [InlineData(true, 4)]
        public void DataReader_XmlTextBudgetCountsReplacedCellRecords(bool replaceWithText, long characters) {
            string replacement = replaceWithText ? "<c r=\"A2\" t=\"str\"><v>x</v></c>" : "<c r=\"A2\"/>";
            string path = CreateXmlTextBudgetWorkbook(
                "<row r=\"2\"><c r=\"A2\" t=\"str\"><v>abc</v></c>" + replacement + "</row>");
            try {
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path,
                    new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = characters })) {
                    Assert.True(reader.Read());
                    Assert.Equal(replaceWithText ? (object)"x" : DBNull.Value, reader.GetValue(0));
                }
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path,
                    new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = characters - 1 })) {
                    Assert.True(reader.Read());
                    AssertXmlTextBudgetFailure(() => reader.IsDBNull(0));
                    AssertXmlTextBudgetReaderIsTerminal(reader);
                }
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void DataReader_XmlTextBudgetRejectsOversizedHeaderDuringOpen() {
            string path = CreateXmlTextBudgetWorkbook(
                "<row r=\"2\"><c r=\"A2\"><v>1</v></c></row>", columns: 2,
                headerCells: "<c r=\"A1\" t=\"inlineStr\"><is><t>One</t></is></c>"
                    + "<c r=\"B1\" t=\"inlineStr\"><is><t>Two</t></is></c>");
            try {
                AssertXmlTextBudgetFailure(() => {
                    using DbDataReader reader = ExcelDocument.OpenDataReader(path,
                        new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = 5 });
                });
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void DataReader_XmlTextBudgetSharesOneAllowanceForBufferedAndRepeatedPhysicalRows(bool repeatedRow) {
            string rows = repeatedRow
                ? "<row r=\"2\"><c r=\"A2\" t=\"str\"><v>abc</v></c></row>"
                    + "<row r=\"2\"><c r=\"A2\" t=\"str\"><v>def</v></c></row>"
                : "<row r=\"3\"><c r=\"A3\" t=\"str\"><v>abc</v></c></row>"
                    + "<row r=\"2\"><c r=\"A2\" t=\"str\"><v>def</v></c></row>";
            string path = CreateXmlTextBudgetWorkbook(rows);
            try {
                // Header (1), both physical values (3 + 3), and final row value (1) total eight.
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path,
                    new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = 8 })) {
                    Assert.True(reader.Read());
                    Assert.Equal("def", reader.GetString(0));
                    Assert.True(reader.Read());
                    if (repeatedRow) Assert.True(reader.IsDBNull(0));
                    else Assert.Equal("abc", reader.GetString(0));
                }
                AssertXmlTextBudgetFailure(() => {
                    using DbDataReader reader = ExcelDocument.OpenDataReader(path,
                        new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = 7 });
                });
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void DataReader_XmlTextBudgetPreservesCancellationWithoutPublishingPartialValues() {
            string path = CreateXmlTextBudgetWorkbook(
                "<row r=\"2\"><c r=\"A2\"><v>1</v></c>"
                + "<c r=\"B2\" t=\"str\"><v>abc</v></c></row>", columns: 2);
            try {
                using var cancellation = new CancellationTokenSource();
                using DbDataReader reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions {
                    MaxXmlDataReaderBufferedCharacters = 4,
                    CancellationToken = cancellation.Token,
                    CellValueConverter = context => {
                        if (context.RawText == "abc") cancellation.Cancel();
                        return ExcelCellValue.NotHandled;
                    }
                });
                Assert.True(reader.Read());
                Assert.ThrowsAny<OperationCanceledException>(() => reader.GetInt32(0));
                Assert.ThrowsAny<OperationCanceledException>(() => reader.GetValue(0));
                Assert.ThrowsAny<OperationCanceledException>(() => reader.IsDBNull(0));
                Assert.ThrowsAny<OperationCanceledException>(() => reader.GetValues(new object[2]));
                reader.Close();
                Assert.True(reader.IsClosed);
            } finally {
                File.Delete(path);
            }
        }

        private static string CreateXmlTextBudgetWorkbook(string rows, int columns = 1, string? headerCells = null,
            bool multipleSheets = false) {
            string path = CreateCompactFastPathWorkbook();
            try {
                if (multipleSheets) {
                    using var document = ExcelDocument.Load(path);
                    document.AddWorksheet("Other").CellValue(1, 1, 1);
                    document.Save();
                }
                string headers = headerCells ?? string.Concat(Enumerable.Range(0, columns).Select(column =>
                    $"<c r=\"{(char)('A' + column)}1\" t=\"inlineStr\"><is><t>{(char)('A' + column)}</t></is></c>"));
                string xml = "<?xml version=\"1.0\" encoding=\"utf-16\"?>"
                    + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                    + $"<dimension ref=\"A1:{(char)('A' + columns - 1)}4097\"/><sheetData>"
                    + "<row r=\"1\">" + headers + "</row>" + rows
                    + "<row r=\"4097\"><c r=\"A4097\"><v>0</v></c></row></sheetData></worksheet>";
                // UTF16 prevents indexed UTF8 selection; a real later row selects XML streaming over SDK chunks.
                byte[] bytes = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray();
                ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", bytes);
                return path;
            } catch {
                File.Delete(path);
                throw;
            }
        }

        private static void ReadXmlTextBudgetFirstValue(DbDataReader reader, string firstAccess) {
            if (firstAccess == "scalar") reader.GetInt32(0);
            else if (firstAccess == "value") reader.GetValue(0);
            else if (firstAccess == "null") reader.IsDBNull(0);
            else reader.GetValues(new object[reader.FieldCount]);
        }

        private static void AssertXmlTextBudgetFailure(Action action) {
            InvalidDataException error = Assert.Throws<InvalidDataException>(action);
            Assert.Contains(nameof(ExcelReadOptions.MaxXmlDataReaderBufferedCharacters), error.Message);
        }

        private static void AssertXmlTextBudgetReaderIsTerminal(DbDataReader reader) {
            Assert.False(reader.IsClosed);
            AssertXmlTextBudgetFailure(() => reader.Read());
            AssertXmlTextBudgetFailure(() => reader.GetInt32(0));
            AssertXmlTextBudgetFailure(() => reader.GetValue(0));
            AssertXmlTextBudgetFailure(() => reader.IsDBNull(0));
            AssertXmlTextBudgetFailure(() => reader.GetValues(new object[reader.FieldCount]));
            reader.Close();
            Assert.True(reader.IsClosed);
            Assert.Throws<InvalidOperationException>(() => reader.Read());
        }
    }
}
