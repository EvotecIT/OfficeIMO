using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    public static IEnumerable<object[]> TypedFormulaFallbackCases() {
        foreach (int lastRow in new[] { 3, 5000 })
            foreach (bool useCachedResult in new[] { false, true })
                foreach (string api in new[] { "Automatic", "Sequential", "Parallel", "Stream" })
                    yield return new object[] { lastRow, useCachedResult, api };
    }

    [Theory]
    [InlineData("Style", ExcelExecutionMode.Automatic)]
    [InlineData("Style", ExcelExecutionMode.Sequential)]
    [InlineData("SharedString", ExcelExecutionMode.Automatic)]
    [InlineData("SharedString", ExcelExecutionMode.Sequential)]
    public void Reader_TypedIndexFallbackDoesNotHideInvalidCellReferences(string invalidReference, ExcelExecutionMode mode) {
        string path = CreateCompactFastPathWorkbook();
        string cell = invalidReference == "Style"
            ? "<c r=\"A2\" s=\"99999\"><v>42</v></c>"
            : "<c r=\"A2\" t=\"s\"><v>99999</v></c>";
        string xml = $$"""
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A2"/><sheetData>
              <row r="1"><c r="A1" t="str"><v>Id</v></c></row><row r="2">{{cell}}</row>
            </sheetData></worksheet>
            """;
        try {
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            using var owner = ExcelDocumentReader.Open(path);
            var exception = Assert.Throws<InvalidDataException>(() => owner.GetSheet("Data")
                .ReadObjects<TypedFormulaFallbackRow>("A1:A2", mode).ToArray());
            Assert.Contains("A2", exception.Message);
            Assert.Contains(invalidReference == "Style" ? "style" : "shared string", exception.Message);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [MemberData(nameof(TypedFormulaFallbackCases))]
    public void Reader_TypedReadersIgnoreUncachedSharedFormulaOutsideRequestedColumns(
        int lastRow, bool useCachedResult, string api) {
        string path = CreateCompactFastPathWorkbook();
        string finalRow = lastRow > 3 ? $"<row r=\"{lastRow}\"><c r=\"A{lastRow}\"><v>46</v></c></row>" : string.Empty;
        string xml = $$"""
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B{{lastRow}}"/><sheetData>
              <row r="1"><c r="A1" t="str"><v>Id</v></c></row>
              <row r="2"><c r="A2"><v>42</v></c><c r="B2"><f t="shared" si="0" ref="B2:B3">A2+1</f><v>43</v></c></row>
              <row r="3"><c r="A3"><v>44</v></c><c r="B3"><f t="shared" si="0"/></c></row>
              {{finalRow}}
            </sheetData></worksheet>
            """;
        try {
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            using var owner = ExcelDocumentReader.Open(path, new ExcelReadOptions { UseCachedFormulaResult = useCachedResult });
            var sheet = owner.GetSheet("Data");
            string range = $"A1:A{lastRow}";
            var rows = (api == "Stream"
                ? sheet.ReadObjectsStream<TypedFormulaFallbackRow>(range)
                : sheet.ReadObjects<TypedFormulaFallbackRow>(range, (ExcelExecutionMode)Enum.Parse(typeof(ExcelExecutionMode), api))).ToArray();
            Assert.Equal(lastRow - 1, rows.Length);
            Assert.Equal(42, rows[0].Id);
            Assert.Equal(44, rows[1].Id);
            Assert.Equal(1, rows[0].AssignmentCount);
            Assert.Equal(1, rows[1].AssignmentCount);
            if (lastRow > 3) {
                Assert.Equal(0, rows[2].Id);
                Assert.Equal(0, rows[2].AssignmentCount);
                Assert.Equal(46, rows[rows.Length - 1].Id);
            }
        } finally {
            File.Delete(path);
        }
    }

    public sealed class TypedFormulaFallbackRow {
        private int _id;
        public int Id {
            get => _id;
            set { _id = value; AssignmentCount++; }
        }
        public int AssignmentCount { get; private set; }
    }
}
