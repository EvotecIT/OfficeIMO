using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData("http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [InlineData("http://purl.oclc.org/ooxml/spreadsheetml/main")]
    public void OpenDataReader_SharedStrings_UseOnlyOwnedVisibleText(string spreadsheetNamespace) {
        string[] expected = { "shown", " A<&B&🐢 \t ", " standalone <&> ", string.Empty };
        const string items = """
            <s:si>
              <s:r><s:t>shown</s:t><x:t>hidden-in-run</x:t></s:r>
              <x:r><x:t>hidden</x:t><s:t>also-hidden</s:t></x:r>
              <x:t>hidden-direct</x:t>
              <x:payload><s:r><s:t>nested-hidden</s:t></s:r></x:payload>
              <s:rPh sb="0" eb="5"><s:t>phonetic</s:t></s:rPh>
            </s:si>
            <s:si><s:r><s:t xml:space="preserve"> A<![CDATA[<&]]>B&amp;🐢</s:t></s:r><s:r><s:t xml:space="preserve"> 	 </s:t></s:r></s:si>
            <s:si><s:t xml:space="preserve"> standalone &lt;&amp;&gt; </s:t></s:si>
            <s:si><s:r><s:t/></s:r><s:r><s:t></s:t></s:r></s:si>
            """;
        string path = CreateSharedStringOwnershipWorkbook(items, spreadsheetNamespace, expected.Length);
        try {
            AssertSharedStringReaderValues(path, expected);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [InlineData("http://purl.oclc.org/ooxml/spreadsheetml/main")]
    public void OpenDataReader_SharedStrings_KeepIndexesAndLimitsForDirectTableItems(string spreadsheetNamespace) {
        const string items = """
            <x:si><x:t>foreign-before</x:t></x:si>
            <x:payload><s:si><s:t>nested-before</s:t></s:si></x:payload>
            <s:si><s:t>first</s:t><x:payload><s:si><s:t>nested-inside</s:t></s:si></x:payload></s:si>
            <x:si><s:t>foreign-between</s:t></x:si>
            <s:si><s:t>second</s:t></s:si>
            <s:extLst><s:ext uri="urn:shared-string-ownership"><x:payload><s:si><s:t>nested-after</s:t></s:si></x:payload></s:ext></s:extLst>
            """;
        string path = CreateSharedStringOwnershipWorkbook(items, spreadsheetNamespace, itemCount: 2);
        try {
            AssertSharedStringReaderValues(path, new[] { "first", "second" });
        } finally {
            File.Delete(path);
        }
    }

    private static void AssertSharedStringReaderValues(string path, string[] expected) {
        var options = new ExcelReadOptions {
            MaxSharedStringItems = expected.Length,
            MaxSharedStringItemCharacters = expected.Max(value => value.Length),
            MaxSharedStringCharacters = expected.Sum(value => (long)value.Length)
        };
        foreach (CultureInfo culture in new[] { CultureInfo.InvariantCulture, CultureInfo.GetCultureInfo("de-DE") }) {
            options.Culture = culture;
            using (var reader = ExcelDocument.OpenDataReader(path, options)) {
                Assert.Equal("Value", reader.GetName(0));
                for (int row = 0; row < expected.Length; row++) {
                    Assert.True(reader.Read());
                    Assert.Equal(expected[row], reader.GetString(0));
                }
                Assert.False(reader.Read());
            }
        }

        // The SDK-backed range helper is additional internal compatibility coverage.
        using var owner = ExcelDocumentReader.Open(path, options);
        object?[,] values = owner.GetSheet("Data").ReadRange(
            $"A2:A{expected.Length + 1}", ExcelExecutionMode.Sequential);
        Assert.Equal(expected.Length, values.GetLength(0));
        for (int row = 0; row < expected.Length; row++) Assert.Equal(expected[row], values[row, 0]);
    }

    private static string CreateSharedStringOwnershipWorkbook(string items, string spreadsheetNamespace, int itemCount) {
        string path = CreateCompactFastPathWorkbook();
        using (var package = SpreadsheetDocument.Open(path, true)) {
            var part = package.WorkbookPart!.SharedStringTablePart
                ?? package.WorkbookPart.AddNewPart<SharedStringTablePart>();
            part.SharedStringTable = new SharedStringTable();
            part.SharedStringTable.Save();
        }

        string table = $$"""
            <s:sst xmlns:s="{{spreadsheetNamespace}}" xmlns:x="urn:shared-string-ownership"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x"
                count="{{itemCount}}" uniqueCount="{{itemCount}}">{{items}}</s:sst>
            """;
        ReplaceZipEntry(path, "xl/sharedStrings.xml", Encoding.UTF8.GetBytes(table));

        string rows = string.Concat(Enumerable.Range(0, itemCount).Select(index =>
            $"<row r=\"{index + 2}\"><c r=\"A{index + 2}\" t=\"s\"><v>{index}</v></c></row>"));
        string worksheet = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + $"<dimension ref=\"A1:A{itemCount + 1}\"/><sheetData>"
            + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Value</t></is></c></row>"
            + rows + "</sheetData></worksheet>";
        ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(worksheet));
        return path;
    }
}
