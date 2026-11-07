using DocumentFormat.OpenXml.Features;
using DocumentFormat.OpenXml.Packaging;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
#if NET8_0_OR_GREATER
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StylesValidationObservesCancellation(bool loadedSdk) {
        string styles = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + "<cellXfs count=\"1\"><xf numFmtId=\"0\"/></cellXfs><cellStyles>"
            + string.Concat(Enumerable.Repeat("<cellStyle name=\"Default\" xfId=\"0\"/>", 1000))
            + "</cellStyles></styleSheet>";
        using var cancellation = new System.Threading.CancellationTokenSource();
        if (loadedSdk) {
            string path = CreateCompactFastPathWorkbook();
            try {
                ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes(styles));
                using var document = SpreadsheetDocument.Open(path, false, new OpenSettings {
                    MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(
                        MarkupCompatibilityProcessMode.ProcessLoadedPartsOnly, DocumentFormat.OpenXml.FileFormatVersions.Office2010)
                });
                var part = document.WorkbookPart!.WorkbookStylesPart!;
                document.AddPartRootEventsFeature();
                bool loaded = false;
                document.Features.Get<IPartRootEventsFeature>()!.Change += change => {
                    if (change.Type == EventType.Created && ReferenceEquals(change.Argument, part)) {
                        loaded = true;
                        cancellation.Cancel();
                    }
                };
                var provider = new StylesCacheProvider(document);
                Exception? error = Record.Exception(() => { _ = provider.GetValue(cancellation.Token); });
                Assert.True(loaded);
                Assert.True(cancellation.IsCancellationRequested);
                var cancelled = Assert.IsType<OperationCanceledException>(error);
                Assert.Equal(cancellation.Token, cancelled.CancellationToken);
            } finally { File.Delete(path); }
            return;
        }
        using var tail = new TailCancellationInput(Encoding.UTF8.GetBytes(styles), 4096, cancellation) { Armed = true };
        var streamProvider = new StylesCacheProvider(() => tail);
        var exception = Assert.Throws<OperationCanceledException>(() => streamProvider.GetValue(cancellation.Token));
        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(0, tail.ReadsAfterCancellation);
    }
#endif

    [Theory]
    [InlineData(false, "trailing-markup")]
    [InlineData(true, "trailing-markup")]
    [InlineData(false, "entity")]
    [InlineData(true, "entity")]
    [InlineData(false, "character")]
    [InlineData(true, "character")]
    [InlineData(false, "duplicate-attribute")]
    [InlineData(true, "duplicate-attribute")]
    [InlineData(false, "sdk-trailing-markup")]
    [InlineData(true, "sdk-trailing-markup")]
    public void StylesValidationRejectsMalformedContentAfterCellFormats(bool streamedPart, string damage) {
        string path = CreateCompactFastPathWorkbook();
        try {
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes("""
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
                  <row r="1"><c r="A1" t="inlineStr"><is><t>Value</t></is></c></row>
                  <row r="2"><c r="A2" s="0"><v>42</v></c></row>
                  <row r="3"><c r="A3" s="0"><v>43</v></c></row>
                </sheetData></worksheet>
                """));
            string padding = streamedPart ? new string('x', 10_000) : "Default";
            string styles = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<cellXfs count=\"1\"><xf numFmtId=\"0\"/></cellXfs>"
                + "<cellStyles count=\"1\"><cellStyle name=\"" + padding
                + "\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>";
            ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes(styles));
            AssertCompactNumericRows(path);
            if (damage == "sdk-trailing-markup") {
                using var validDocument = OpenSdkStylesDocument();
                Assert.Equal(1, new StylesCacheProvider(validDocument).GetValue(default).CellFormatCount);
            }

            string malformed = damage switch {
                "trailing-markup" => styles + "<",
                "sdk-trailing-markup" => styles + "<!--tail--><",
                "entity" => styles.Replace("name=\"" + padding + "\"", "name=\"" + padding + "&missing;\""),
                "character" => styles.Replace("name=\"" + padding + "\"", "name=\"" + padding + "bad\u0001text\""),
                "duplicate-attribute" => styles.Replace("name=\"" + padding + "\"", "name=\"" + padding + "\" name=\"second\""),
                _ => throw new ArgumentOutOfRangeException(nameof(damage))
            };
            ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes(malformed));
            Assert.Throws<System.Xml.XmlException>(() => {
                if (damage == "sdk-trailing-markup") {
                    using var document = OpenSdkStylesDocument();
                    _ = new StylesCacheProvider(document).GetValue(default);
                    return;
                }
                using var reader = ExcelDocument.OpenReader(path);
                _ = reader.FieldCount;
            });
        } finally { File.Delete(path); }

        SpreadsheetDocument OpenSdkStylesDocument() => SpreadsheetDocument.Open(path, false, new OpenSettings {
            MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(
                MarkupCompatibilityProcessMode.ProcessLoadedPartsOnly, DocumentFormat.OpenXml.FileFormatVersions.Office2010)
        });
    }

    [Theory]
    [InlineData("plain", false)]
    [InlineData("plain", true)]
    [InlineData("entities", false)]
    [InlineData("entities", true)]
    [InlineData("padded-style", false)]
    [InlineData("padded-style", true)]
    [InlineData("long-style", false)]
    [InlineData("long-style", true)]
    [InlineData("unknown-type", false)]
    [InlineData("unknown-type", true)]
    [InlineData("alternate-formats", false)]
    [InlineData("alternate-formats", true)]
    [InlineData("process-number-formats", false)]
    [InlineData("process-number-formats", true)]
    [InlineData("process-cell-formats", false)]
    [InlineData("process-cell-formats", true)]
    [InlineData("process-root", false)]
    [InlineData("process-root", true)]
    [InlineData("process-local-scope", false)]
    [InlineData("process-local-scope", true)]
    [InlineData("process-prefix-scope", false)]
    [InlineData("process-prefix-scope", true)]
    [InlineData("process-wildcard", false)]
    [InlineData("process-wildcard", true)]
    [InlineData("alternate-cell-formats", false)]
    [InlineData("alternate-cell-formats", true)]
    [InlineData("alternate-known-choice", false)]
    [InlineData("alternate-known-choice", true)]
    [InlineData("alternate-first-choice", false)]
    [InlineData("alternate-first-choice", true)]
    [InlineData("alternate-mixed-requires", false)]
    [InlineData("alternate-mixed-requires", true)]
    [InlineData("alternate-root", false)]
    [InlineData("alternate-root", true)]
    [InlineData("alternate-nested", false)]
    [InlineData("alternate-nested", true)]
    public void Reader_XmlMetadataPreservesTypedAndMaterializedValues(string shape, bool date1904) {
        string path = CreateCompactFastPathWorkbook();
        try {
            using (var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, true)) {
                workbook.WorkbookPart!.Workbook.WorkbookProperties =
                    new DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties { Date1904 = date1904 };
            }
            string numberType = shape == "entities" ? "&#110;" : shape == "unknown-type" ? new string('x', 31) + "😀tail" : "n";
            string booleanType = shape == "entities" ? "&#98;" : "b";
            string style = shape switch {
                "entities" => "&#49;",
                "padded-style" => " +1 ",
                "long-style" => new string('0', 80) + "1",
                _ => "1"
            };
            string xml = $$"""
                <x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData>
                  <x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>Id</x:t></x:is></x:c><x:c r="B1" t="inlineStr"><x:is><x:t>Active</x:t></x:is></x:c><x:c r="C1" t="inlineStr"><x:is><x:t>Name</x:t></x:is></x:c><x:c r="D1" t="inlineStr"><x:is><x:t>Created</x:t></x:is></x:c></x:row>
                  <x:row r="2"><x:c r="A2" t="{{numberType}}" s="0"><x:v>42</x:v></x:c><x:c r="B2" t="{{booleanType}}"><x:v>1</x:v></x:c><x:c r="C2" t="str"><x:v>Alpha</x:v></x:c><x:c r="D2" t="n" s="{{style}}"><x:v>45351</x:v></x:c></x:row>
                  <x:row r="3"><x:c r="A3" t="n" s="0"><x:v>43</x:v></x:c><x:c r="B3" t="b"><x:v>0</x:v></x:c><x:c r="C3" t="str"><x:v>Beta</x:v></x:c><x:c r="D3" s="1"><x:v>45352</x:v></x:c></x:row>
                </x:sheetData></x:worksheet>
                """;
            if (shape == "alternate-cell-formats") {
                xml = xml.Replace("r=\"A2\" t=\"n\" s=\"0\"", "r=\"A2\" t=\"n\" s=\"2\"")
                    .Replace("r=\"A3\" t=\"n\" s=\"0\"", "r=\"A3\" t=\"n\" s=\"2\"");
            }
            ReplaceXmlMetadataWorksheet(path, xml);
            string styles = shape == "alternate-formats" ? """
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p="urn:private:future-style">
                  <numFmts count="1"><mc:AlternateContent>
                    <mc:Choice Requires="p"><numFmt numFmtId="165" formatCode="0.00"/></mc:Choice>
                    <mc:Fallback><numFmt numFmtId="165" formatCode="yyyy-mm-dd"/></mc:Fallback>
                  </mc:AlternateContent></numFmts>
                  <cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="165" applyNumberFormat="1"/></cellXfs>
                </styleSheet>
                """ : """
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>
                """;
            if (shape.StartsWith("alternate-", StringComparison.Ordinal) && shape != "alternate-formats") {
                const string dateFormat = "<numFmt numFmtId=\"165\" formatCode=\"yyyy-mm-dd\"/>";
                const string numberFormat = "<numFmt numFmtId=\"165\" formatCode=\"0.00\"/>";
                string formats = "<numFmts count=\"1\">" + dateFormat + "</numFmts>";
                string cellFormats = "<cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"165\"/></cellXfs>";
                string content;
                switch (shape) {
                    case "alternate-cell-formats":
                        cellFormats = "<cellXfs count=\"3\"><xf numFmtId=\"0\"/><mc:AlternateContent>"
                            + "<mc:Choice Requires=\"p\"><xf numFmtId=\"0\"/></mc:Choice>"
                            + "<mc:Fallback><xf numFmtId=\"165\"/></mc:Fallback></mc:AlternateContent><xf numFmtId=\"0\"/></cellXfs>";
                        content = formats + cellFormats;
                        break;
                    case "alternate-known-choice":
                        content = "<numFmts count=\"1\"><mc:AlternateContent><mc:Choice Requires=\"s\">" + dateFormat
                            + "</mc:Choice><mc:Fallback>" + numberFormat + "</mc:Fallback></mc:AlternateContent></numFmts>" + cellFormats;
                        break;
                    case "alternate-first-choice":
                        content = "<numFmts count=\"1\"><mc:AlternateContent><mc:Choice Requires=\"p\">" + numberFormat
                            + "</mc:Choice><mc:Choice Requires=\"s\">" + dateFormat + "</mc:Choice><mc:Choice Requires=\"s\">" + numberFormat
                            + "</mc:Choice><mc:Fallback>" + numberFormat + "</mc:Fallback></mc:AlternateContent></numFmts>" + cellFormats;
                        break;
                    case "alternate-mixed-requires":
                        content = "<numFmts count=\"1\"><mc:AlternateContent><mc:Choice Requires=\"s p\">" + numberFormat
                            + "</mc:Choice><mc:Fallback>" + dateFormat + "</mc:Fallback></mc:AlternateContent></numFmts>" + cellFormats;
                        break;
                    case "alternate-root":
                        content = "<mc:AlternateContent><mc:Choice Requires=\"p\"><numFmts count=\"1\">" + numberFormat
                            + "</numFmts>" + cellFormats + "</mc:Choice><mc:Fallback>" + formats + cellFormats + "</mc:Fallback></mc:AlternateContent>";
                        break;
                    case "alternate-nested":
                        content = "<mc:AlternateContent><mc:Choice Requires=\"s\"><p:payload><numFmts count=\"1\"><mc:AlternateContent>"
                            + "<mc:Choice Requires=\"p\">" + numberFormat + "</mc:Choice><mc:Fallback>" + dateFormat
                            + "</mc:Fallback></mc:AlternateContent></numFmts><cellXfs count=\"2\"><xf numFmtId=\"0\"/><mc:AlternateContent>"
                            + "<mc:Choice Requires=\"s\"><xf numFmtId=\"165\"/></mc:Choice><mc:Fallback><xf numFmtId=\"0\"/></mc:Fallback>"
                            + "</mc:AlternateContent></cellXfs></p:payload></mc:Choice><mc:Fallback><cellXfs count=\"2\"><xf numFmtId=\"0\"/>"
                            + "<xf numFmtId=\"0\"/></cellXfs></mc:Fallback></mc:AlternateContent>";
                        break;
                    default: throw new ArgumentOutOfRangeException(nameof(shape));
                }
                styles = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                    + " xmlns:s=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:p=\"urn:private:future-style\""
                    + " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"p\" mc:ProcessContent=\"p:payload\">"
                    + content + "</styleSheet>";
            }
            if (shape.StartsWith("process-", StringComparison.Ordinal)) {
                string formats = "<numFmts count=\"1\"><numFmt numFmtId=\"165\" formatCode=\"yyyy-mm-dd\"/></numFmts>";
                string cellFormats = "<cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"165\"/></cellXfs>";
                string processContent = "p:payload";
                switch (shape) {
                    case "process-number-formats":
                        formats = formats.Replace("<numFmt numFmtId", "<p:payload><numFmt numFmtId").Replace("</numFmts>", "</p:payload></numFmts>");
                        break;
                    case "process-cell-formats":
                        cellFormats = cellFormats.Replace("<xf numFmtId=\"165\"/>", "<p:payload><xf numFmtId=\"165\"/></p:payload>");
                        break;
                    case "process-root":
                        formats = "<p:payload>" + formats;
                        cellFormats += "</p:payload>";
                        break;
                    case "process-local-scope":
                        processContent = "";
                        formats = formats.Replace("<numFmts count=\"1\">", "<numFmts count=\"1\" mc:ProcessContent=\"p:payload\"><p:payload>")
                            .Replace("</numFmts>", "</p:payload></numFmts>");
                        cellFormats = cellFormats.Replace("<xf numFmtId=\"165\"/>", "<p:payload><xf numFmtId=\"0\"/></p:payload><xf numFmtId=\"165\"/>");
                        break;
                    case "process-prefix-scope":
                        formats = formats.Replace("<numFmt numFmtId", "<p:payload><numFmt numFmtId").Replace("</numFmts>", "</p:payload></numFmts>");
                        cellFormats = cellFormats.Replace("<cellXfs count=\"2\">", "<cellXfs count=\"2\" xmlns:p=\"urn:private:other-style\" mc:Ignorable=\"p\">")
                            .Replace("<xf numFmtId=\"165\"/>", "<p:payload><xf numFmtId=\"0\"/></p:payload><xf numFmtId=\"165\"/>");
                        break;
                    case "process-wildcard":
                        processContent = "p:*";
                        formats = "<p:numberFormats>" + formats + "</p:numberFormats>";
                        cellFormats = "<p:cellFormats>" + cellFormats + "</p:cellFormats>";
                        break;
                }
                styles = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                    + " xmlns:p=\"urn:private:future-style\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\""
                    + " mc:Ignorable=\"p\" mc:ProcessContent=\"" + processContent + "\">" + formats + cellFormats + "</styleSheet>";
            }
            ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes(styles));
            var options = new ExcelReaderOptions { NumericAsDecimal = true, InferDataTableColumnTypes = false };
            object[][] expected = [
                [42m, true, "Alpha", new DateTime(2024, 2, 29).AddDays(date1904 ? 1462 : 0)],
                [43m, false, "Beta", new DateTime(2024, 3, 1).AddDays(date1904 ? 1462 : 0)]
            ];
            using (var reader = ExcelDocument.OpenReader(path, options)) {
                for (int row = 0; row < 2; row++) {
                    Assert.True(reader.Read());
                    Assert.Equal(expected[row][0], reader.GetDecimal(0));
                    Assert.Equal(expected[row][1], reader.GetBoolean(1));
                    Assert.Equal(expected[row][2], reader.GetString(2));
                    Assert.Equal(expected[row][3], reader.GetDateTime(3));
                    for (int column = 0; column < 4; column++) Assert.Equal(expected[row][column], reader.GetValue(column));
                }
                Assert.False(reader.Read());
            }
            using var owner = ExcelDocumentReader.Open(path, options);
            var sheet = owner.GetSheet("Data");
            var values = sheet.ReadRange("A2:D3", ExcelExecutionMode.Sequential);
            using var table = sheet.ReadRangeAsDataTable("A1:D3", headersInFirstRow: true);
            var objects = sheet.ReadObjects<XmlMetadataRecord>("A1:D3").ToArray();
            Assert.Equal(2, table.Rows.Count);
            Assert.Equal(2, objects.Length);
            for (int row = 0; row < 2; row++) {
                for (int column = 0; column < 4; column++) {
                    Assert.Equal(expected[row][column], values[row, column]);
                    Assert.Equal(expected[row][column], table.Rows[row][column]);
                }
                Assert.Equal(expected[row][0], objects[row].Id);
                Assert.Equal(expected[row][1], objects[row].Active);
                Assert.Equal(expected[row][2], objects[row].Name);
                Assert.Equal(expected[row][3], objects[row].Created);
            }
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(OfficeOpenXmlFileFormatVersion.Office2007, true, false)]
    [InlineData(OfficeOpenXmlFileFormatVersion.Office2010, true, true)]
    [InlineData(OfficeOpenXmlFileFormatVersion.Office2007, false, true)]
    [InlineData(OfficeOpenXmlFileFormatVersion.Office2010, false, false)]
    public void LoadedReader_UsesConfiguredStyleCompatibilityVersion(OfficeOpenXmlFileFormatVersion version, bool choiceDate, bool dateStyle) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string choiceFormat = choiceDate ? "yyyy-mm-dd" : "0.00";
            string fallbackFormat = choiceDate ? "0.00" : "yyyy-mm-dd";
            ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes($$"""
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
                    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14">
                  <numFmts count="1"><mc:AlternateContent>
                    <mc:Choice Requires="x14"><numFmt numFmtId="165" formatCode="{{choiceFormat}}"/></mc:Choice>
                    <mc:Fallback><numFmt numFmtId="165" formatCode="{{fallbackFormat}}"/></mc:Fallback>
                  </mc:AlternateContent></numFmts>
                  <cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="165"/></cellXfs>
                </styleSheet>
                """));
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes("""
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
                  <row><c t="inlineStr"><is><t>Value</t></is></c></row>
                  <row><c s="1"><v>45351</v></c></row>
                </sheetData></worksheet>
                """));
            using var document = ExcelDocument.Load(path, new ExcelLoadOptions {
                AccessMode = DocumentAccessMode.ReadOnly,
                OpenSettings = new OfficeOpenXmlLoadSettings {
                    MarkupCompatibilityMode = OfficeOpenXmlMarkupCompatibilityMode.ProcessAllParts,
                    MarkupCompatibilityTargetVersion = version
                }
            });
            using var reader = document.CreateDataReader(new ExcelReaderOptions { NumericAsDecimal = false });
            Assert.True(reader.Read());
            if (dateStyle) Assert.Equal(new DateTime(2024, 2, 29), Assert.IsType<DateTime>(reader.GetValue(0)));
            else Assert.Equal(45351d, Assert.IsType<double>(reader.GetValue(0)));
            Assert.False(reader.Read());
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("")]
    [InlineData("4294967296")]
    [InlineData("-1")]
    [InlineData("long")]
    [InlineData("surrogate")]
    public void DataReader_XmlMetadataRejectsInvalidStylesWithOriginalCellReference(string style) {
        string path = CreateCompactFastPathWorkbook();
        try {
            if (style == "long") style = new string('0', 80) + "x";
            if (style == "surrogate") style = new string('0', 31) + "😀0";
            ReplaceXmlMetadataWorksheet(path, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                + $"<row r=\"2\"><c r=\"&#67;2\" t=\"n\" s=\"{style}\"><v>42</v></c></row></sheetData></worksheet>");
            var error = Assert.Throws<InvalidDataException>(() => {
                using var reader = ExcelDocument.OpenReader(path);
 _ = reader.FieldCount;             });
            Assert.Contains("cell C2 contains an invalid cell style index", error.Message);
        } finally { File.Delete(path); }
    }

    private static void ReplaceXmlMetadataWorksheet(string path, string xml) =>
        ReplaceZipEntry(path, "xl/worksheets/sheet1.xml",
            Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray());

    private sealed class XmlMetadataRecord {
        public decimal Id { get; set; }
        public bool Active { get; set; }
        public string? Name { get; set; }
        public DateTime Created { get; set; }
    }
}
