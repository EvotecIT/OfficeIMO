using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    public static IEnumerable<object[]> TypedPartialRowUpdateCases() {
        foreach (int lastRow in new[] { 5, 5000 })
            foreach (bool utf16 in new[] { false, true })
                foreach (string path in new[] { "Default", "Converter", "Dom" })
                    foreach (string api in new[] { "Stream", "Automatic", "Sequential", "Parallel", "DataReader" })
                        yield return new object[] { lastRow, utf16, path, api };
    }

    [Theory]
    [MemberData(nameof(TypedPartialRowUpdateCases))]
    public void Reader_TypedReadersMergePartialRowUpdates(int lastRow, bool utf16, string readerPath, string api) {
        string path = CreateCompactFastPathWorkbook();
        string encodingName = utf16 ? "utf-16" : "utf-8";
        string xml = $$"""
            <?xml version="1.0" encoding="{{encodingName}}"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
              <row r="2"><c r="B2" t="str"><v>Id</v></c><c r="C2" t="str"><v>Name</v></c><c r="D2" t="str"><v>Note</v></c><c r="E2" t="str"><v>Count</v></c><c r="F2" t="str"><v>Amount</v></c><c r="G2" t="str"><v>Active</v></c><c r="H2" t="str"><v>Created</v></c></row>
              <row r="3"><c r="B3"><v>1</v></c><c r="C3" t="str"><v>first</v></c><c r="D3" t="str"><v>old note</v></c><c r="E3"><v>9</v></c><c r="F3"><v>2</v></c><c r="G3" t="b"><v>1</v></c><c r="H3" t="d"><v>2020-01-02T00:00:00</v></c></row>
              <row r="5"><c r="B5"><v>2</v></c><c r="C5" t="str"><v>second</v></c><c r="D5" t="str"><v>retained note</v></c></row>
              <row r="{{lastRow + 1}}"><c r="A{{lastRow + 1}}" t="str"><v>outside</v></c></row>
              <row r="3"><c r="C3" t="str"><v>patched</v></c><c r="D3" t="str"><v/></c><c r="E3" t="str"><v/></c><c r="F3" t="str"><v/></c><c r="G3" t="str"><v/></c><c r="H3" t="str"><v/></c></row>
              <row r="5"><c r="C5" t="str"><v>changed second</v></c></row>
            </sheetData></worksheet>
            """;
        try {
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", (utf16 ? Encoding.Unicode : Encoding.UTF8).GetBytes(xml));
            var options = new ExcelReadOptions { MaxPendingTypedRows = 2 };
            if (readerPath == "Converter") options.CellValueConverter = static _ => ExcelCellValue.NotHandled;
            if (readerPath == "Dom") options.Culture = CultureInfo.GetCultureInfo("fr-FR");
            using var document = readerPath == "Dom" ? SpreadsheetDocument.Open(path, true) : null;
            using var owner = document == null
                ? ExcelDocumentReader.Open(path, options)
                : ExcelDocumentReader.Wrap(document, options);
            var sheet = owner.GetSheet("Data");
            string range = $"B2:H{lastRow}";
            if (api == "DataReader") {
                using var reader = sheet.ReadRangeAsDataReader(range, chunkRows: 1, schemaSampleRows: 0);
                int dataRow = 3;
                while (reader.Read()) {
                    if (dataRow == 3) {
                        Assert.Equal(1, reader.GetInt32(0));
                        Assert.Equal("patched", reader.GetString(1));
                        for (int column = 2; column < 7; column++) Assert.Equal(string.Empty, reader.GetValue(column));
                    } else if (dataRow == 5) {
                        Assert.Equal(2, reader.GetInt32(0));
                        Assert.Equal("changed second", reader.GetString(1));
                        Assert.Equal("retained note", reader.GetString(2));
                        for (int column = 3; column < 7; column++) Assert.True(reader.IsDBNull(column));
                    } else {
                        for (int column = 0; column < 7; column++) Assert.True(reader.IsDBNull(column));
                    }
                    dataRow++;
                }
                Assert.Equal(lastRow + 1, dataRow);
                return;
            }
            var rows = (api == "Stream"
                ? sheet.ReadObjectsStream<TypedPartialRowUpdate>(range)
                : sheet.ReadObjects<TypedPartialRowUpdate>(range, (ExcelExecutionMode)Enum.Parse(typeof(ExcelExecutionMode), api))).ToArray();
            Assert.Equal(lastRow - 2, rows.Length);
            Assert.Equal(1, rows[0].Id);
            Assert.Equal("patched", rows[0].Name);
            Assert.Equal(string.Empty, rows[0].Note);
            Assert.Equal(42, rows[0].Count);
            Assert.Equal(0m, rows[0].Amount);
            Assert.False(rows[0].Active);
            Assert.Equal(default, rows[0].Created);
            Assert.Equal(0, rows[1].Id);
            Assert.Null(rows[1].Name);
            Assert.Equal(2, rows[2].Id);
            Assert.Equal("changed second", rows[2].Name);
            Assert.Equal("retained note", rows[2].Note);
            for (int row = 3; row < rows.Length; row++) {
                Assert.Equal(0, rows[row].Id);
                Assert.Null(rows[row].Name);
                Assert.Null(rows[row].Note);
            }
        } finally {
            File.Delete(path);
        }
    }

    public sealed class TypedPartialRowUpdate {
        public int Id { get; set; }
        public string? Name { get; set; }
        public string? Note { get; set; }
        public int Count { get; set; } = 42;
        public decimal Amount { get; set; }
        public bool Active { get; set; }
        public DateTime Created { get; set; }
    }

    public static IEnumerable<object[]> TypedFinalRowValueCases() {
        foreach (bool dom in new[] { false, true })
            foreach (string api in new[] { "Stream", "Automatic", "Sequential", "Parallel" })
                yield return new object[] { dom, api };
    }

    [Theory]
    [MemberData(nameof(TypedFinalRowValueCases))]
    public void Reader_TypedReadersApplyHandledNullToInitializedProperty(bool dom, string api) {
        foreach (bool repeatedRow in new[] { false, true }) {
            string path = CreateFinalTypedRowWorkbook(repeatedRow);
            try {
                var options = new ExcelReadOptions {
                    CellValueConverter = static context => context.RawText == string.Empty
                        ? new ExcelCellValue(null)
                        : ExcelCellValue.NotHandled
                };
                var decisions = new List<ExcelExecutionMode>();
                options.Execution.OnDecision = (_, _, mode) => decisions.Add(mode);
                using var document = dom ? SpreadsheetDocument.Open(path, true) : null;
                using var owner = document == null ? ExcelDocumentReader.Open(path, options) : ExcelDocumentReader.Wrap(document, options);
                var sheet = owner.GetSheet("Data");
                var rows = (api == "Stream" ? sheet.ReadObjectsStream<TypedInitializedNote>("A1:B3")
                    : sheet.ReadObjects<TypedInitializedNote>("A1:B3", (ExcelExecutionMode)Enum.Parse(typeof(ExcelExecutionMode), api))).ToArray();
                Assert.Equal(2, rows.Length);
                Assert.Equal(1, rows[0].Id);
                Assert.Null(rows[0].Note);
                Assert.Equal(2, rows[1].Id);
                Assert.Equal("seed", rows[1].Note);
                if (api != "Stream") {
                    Assert.Equal(api == "Parallel" && !repeatedRow ? ExcelExecutionMode.Parallel : ExcelExecutionMode.Sequential,
                        Assert.Single(decisions));
                }
            } finally {
                File.Delete(path);
            }
        }
    }

    [Theory]
    [MemberData(nameof(TypedFinalRowValueCases))]
    public void Reader_TypedReadersDoNotConvertSupersededRowValues(bool dom, string api) {
        string path = CreateFinalTypedRowWorkbook(repeatedRow: true);
        try {
            var options = new ExcelReadOptions {
                CellValueConverter = static context => context.RawText == "superseded"
                    ? throw new InvalidDataException("The final logical row does not contain this value.")
                    : ExcelCellValue.NotHandled
            };
            using var document = dom ? SpreadsheetDocument.Open(path, true) : null;
            using var owner = document == null ? ExcelDocumentReader.Open(path, options) : ExcelDocumentReader.Wrap(document, options);
            var sheet = owner.GetSheet("Data");
            var rows = (api == "Stream" ? sheet.ReadObjectsStream<TypedInitializedNote>("A1:B3")
                : sheet.ReadObjects<TypedInitializedNote>("A1:B3", (ExcelExecutionMode)Enum.Parse(typeof(ExcelExecutionMode), api))).ToArray();
            Assert.Equal(2, rows.Length);
            Assert.Equal(string.Empty, rows[0].Note);
            Assert.Equal("seed", rows[1].Note);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [MemberData(nameof(TypedFinalRowValueCases))]
    public void Reader_TypedReadersDoNotAssignSupersededRowValues(bool dom, string api) {
        string path = CreateFinalTypedRowWorkbook(repeatedRow: true);
        try {
            using var document = dom ? SpreadsheetDocument.Open(path, true) : null;
            using var owner = document == null ? ExcelDocumentReader.Open(path) : ExcelDocumentReader.Wrap(document);
            var sheet = owner.GetSheet("Data");
            var rows = (api == "Stream" ? sheet.ReadObjectsStream<TypedValidatingNote>("A1:B3")
                : sheet.ReadObjects<TypedValidatingNote>("A1:B3", (ExcelExecutionMode)Enum.Parse(typeof(ExcelExecutionMode), api))).ToArray();
            Assert.Equal(2, rows.Length);
            Assert.Equal(1, rows[0].Id);
            Assert.Equal(string.Empty, rows[0].Note);
        } finally {
            File.Delete(path);
        }
    }

    private static string CreateFinalTypedRowWorkbook(bool repeatedRow) {
        string path = CreateCompactFastPathWorkbook();
        string earlier = repeatedRow ? "<row r=\"2\"><c r=\"A2\"><v>1</v></c><c r=\"B2\" t=\"str\"><v>superseded</v></c></row>" : string.Empty;
        string xml = $$"""
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
              <row r="1"><c r="A1" t="str"><v>Id</v></c><c r="B1" t="str"><v>Note</v></c></row>
              {{earlier}}
              <row r="2"><c r="A2"><v>1</v></c><c r="B2" t="str"><v/></c></row>
              <row r="3"><c r="A3"><v>2</v></c></row>
            </sheetData></worksheet>
            """;
        ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
        return path;
    }

    public sealed class TypedInitializedNote {
        public int Id { get; set; }
        public string? Note { get; set; } = "seed";
        public int Count { get; set; } = 42;
    }

    [Theory]
    [MemberData(nameof(TypedFinalRowValueCases))]
    public void Reader_TypedBufferedConvertersPreserveCellStyles(bool dom, string api) {
        foreach (bool treatDates in new[] { false, true }) {
            foreach (string kind in new[] { "str", "b", "d" }) {
                string path = CreateFinalTypedRowWorkbook(repeatedRow: true);
                try {
                    string value = kind == "str" ? "final" : kind == "b" ? "1" : "2001-02-03T00:00:00";
                    string xml = $$"""
                        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
                          <row r="1"><c r="A1" t="str"><v>Id</v></c><c r="B1" t="str"><v>Note</v></c></row>
                          <row r="2"><c r="A2"><v>1</v></c><c r="B2" t="str"><v>superseded</v></c></row>
                          <row r="2"><c r="B2" s="0" t="{{kind}}"><v>{{value}}</v></c></row>
                          <row r="3"><c r="A3"><v>2</v></c></row>
                        </sheetData></worksheet>
                        """;
                    ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
                    var options = new ExcelReadOptions {
                        TreatDatesUsingNumberFormat = treatDates,
                        CellValueConverter = static context => context.TypeHint is ExcelCellValueType.Boolean
                            or ExcelCellValueType.Date || context.RawText == "final"
                            ? new ExcelCellValue(context.StyleIndex == 0 ? "style-zero" : "missing-style")
                            : ExcelCellValue.NotHandled
                    };
                    using var document = dom ? SpreadsheetDocument.Open(path, true) : null;
                    using var owner = document == null ? ExcelDocumentReader.Open(path, options) : ExcelDocumentReader.Wrap(document, options);
                    var sheet = owner.GetSheet("Data");
                    var rows = (api == "Stream" ? sheet.ReadObjectsStream<TypedInitializedNote>("A1:B3")
                        : sheet.ReadObjects<TypedInitializedNote>("A1:B3", (ExcelExecutionMode)Enum.Parse(typeof(ExcelExecutionMode), api))).ToArray();
                    Assert.Equal(2, rows.Length);
                    Assert.Equal(1, rows[0].Id);
                    Assert.Equal("style-zero", rows[0].Note);
                } finally {
                    File.Delete(path);
                }
            }
        }
    }

    [Theory]
    [MemberData(nameof(TypedFinalRowValueCases))]
    public void Reader_TypedReadersDistinguishPresentBlankFromOmittedCell(bool dom, string api) {
        foreach (int lastRow in new[] { 4, 5000 }) {
            foreach (bool utf16 in new[] { false, true }) {
                string path = CreateCompactFastPathWorkbook();
                try {
                    string encodingName = utf16 ? "utf-16" : "utf-8";
                    string xml = $$"""
                        <?xml version="1.0" encoding="{{encodingName}}"?>
                        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="B2:D{{lastRow}}"/><sheetData>
                          <row r="2"><c r="B2" t="str"><v>Id</v></c><c r="C2" t="str"><v>Note</v></c><c r="D2" t="str"><v>Count</v></c></row>
                          <row r="3"><c r="B3"><v>1</v></c><c r="C3"/><c r="D3"/></row>
                          <row r="4"><c r="B4"><v>2</v></c></row>
                        </sheetData></worksheet>
                        """;
                    ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", (utf16 ? Encoding.Unicode : Encoding.UTF8).GetBytes(xml));
                    using var document = dom ? SpreadsheetDocument.Open(path, true) : null;
                    using var owner = document == null ? ExcelDocumentReader.Open(path) : ExcelDocumentReader.Wrap(document);
                    var sheet = owner.GetSheet("Data");
                    string range = $"B2:D{lastRow}";
                    var rows = (api == "Stream" ? sheet.ReadObjectsStream<TypedInitializedNote>(range)
                        : sheet.ReadObjects<TypedInitializedNote>(range, (ExcelExecutionMode)Enum.Parse(typeof(ExcelExecutionMode), api))).ToArray();
                    Assert.Equal(lastRow - 2, rows.Length);
                    Assert.Equal(1, rows[0].Id);
                    Assert.Null(rows[0].Note);
                    Assert.Equal(42, rows[0].Count);
                    Assert.Equal(2, rows[1].Id);
                    Assert.Equal("seed", rows[1].Note);
                    Assert.Equal(42, rows[1].Count);
                    for (int row = 2; row < rows.Length; row++) {
                        Assert.Equal(0, rows[row].Id);
                        Assert.Equal("seed", rows[row].Note);
                        Assert.Equal(42, rows[row].Count);
                    }
                } finally {
                    File.Delete(path);
                }
            }
        }
    }

    public sealed class TypedValidatingNote {
        private string? _note;
        public int Id { get; set; }
        public string? Note {
            get => _note;
            set => _note = value == "superseded" ? throw new InvalidDataException("Only final row values are valid.") : value;
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Reader_LargeTypedAndDataReadersClearBareBlankUpdates(bool dom, bool fillBlanks) {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string xml = """
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="B2:C5000"/><sheetData>
                  <row r="2"><c r="B2" t="str"><v>Id</v></c><c r="C2" t="str"><v>Note</v></c></row>
                  <row r="3"><c r="B3"><v>1</v></c><c r="C3" t="str"><v>old</v></c></row>
                  <row r="4"><c r="B4"><v>2</v></c></row>
                  <row r="5001"><c r="B5001"><v>outside</v></c></row>
                  <row r="3"><c r="C3"/></row>
                </sheetData></worksheet>
                """;
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            var options = new ExcelReadOptions { FillBlanksInRanges = fillBlanks };
            using var document = dom ? SpreadsheetDocument.Open(path, true) : null;
            using var owner = document == null ? ExcelDocumentReader.Open(path, options) : ExcelDocumentReader.Wrap(document, options);
            var sheet = owner.GetSheet("Data");
            using (var reader = sheet.ReadRangeAsDataReader("B2:C5000", chunkRows: 1, schemaSampleRows: 0)) {
                Assert.True(reader.Read());
                Assert.Equal(1, reader.GetInt32(0));
                Assert.True(reader.IsDBNull(1));
            }
            var rows = sheet.ReadObjects<TypedInitializedNote>("B2:C5000", ExcelExecutionMode.Parallel).ToArray();
            Assert.Equal(4998, rows.Length);
            Assert.Equal(1, rows[0].Id);
            Assert.Null(rows[0].Note);
            Assert.Equal("seed", rows[1].Note);
        } finally {
            File.Delete(path);
        }
    }
}
