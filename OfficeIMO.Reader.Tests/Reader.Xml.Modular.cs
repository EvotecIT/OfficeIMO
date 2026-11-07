using OfficeIMO.Reader;
using OfficeIMO.Reader.Xml;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.Word;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderXmlModularTests {
    [Theory]
    [InlineData(16)]
    [InlineData(17)]
    public void WordMlWarningsRetainTheLimitAndSignalOnlyActualTruncation(int diagnosticCount) {
        string omitted = string.Concat(Enumerable.Range(0, diagnosticCount).Select(index => $"<w:unmapped{index}/>"));
        string xml = "<w:wordDocument xmlns:w=\"http://schemas.microsoft.com/office/word/2003/wordml\">" + omitted +
            "<w:body><w:p><w:r><w:t>Retained reader content</w:t></w:r></w:p></w:body></w:wordDocument>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().AddXmlHandler().Build();
        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(xml), "warnings.xml",
            new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent });
        ReaderChunk chunk = Assert.Single(result.Chunks);
        Assert.Contains("Retained reader content", chunk.Text);
        IReadOnlyList<string> warnings = Assert.IsAssignableFrom<IReadOnlyList<string>>(chunk.Warnings);
        Assert.Equal(16, warnings.Count);
        int retained = diagnosticCount > 16 ? 15 : 16;
        for (int index = 0; index < retained; index++) Assert.Contains($"(unmapped{index})", warnings[index]);
        if (diagnosticCount > 16) {
            Assert.StartsWith("WORD_SOURCE_WARNINGS_TRUNCATED:", warnings[15]);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "content-truncated" &&
                diagnostic.Message.StartsWith("WORD_SOURCE_WARNINGS_TRUNCATED:", StringComparison.Ordinal));
        } else {
            Assert.DoesNotContain(warnings, warning => warning.StartsWith("WORD_SOURCE_WARNINGS_TRUNCATED:", StringComparison.Ordinal));
        }
    }

    [Theory]
    [InlineData(false, ".xml")]
    [InlineData(true, ".xml")]
    [InlineData(false, ".txt")]
    [InlineData(true, ".txt")]
    [InlineData(false, ".xlsx")]
    [InlineData(true, ".xlsx")]
    [InlineData(false, ".xls")]
    [InlineData(true, ".xls")]
    [InlineData(false, ".xlsb")]
    [InlineData(true, ".xlsb")]
    public void PreferContentRoutesNativeXmlSpreadsheetToTheExcelProjection(bool useStream, string extension) {
        string fixture = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents",
            "SpreadsheetXmlCorpus", "excel-geometry.xml");
        string path = Path.Combine(Path.GetTempPath(), "officeimo-spreadsheetxml-" + Guid.NewGuid().ToString("N") + extension);
        File.Copy(fixture, path);
        try {
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddExcelHandler().AddXmlHandler().Build();
            var options = new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent };
            using var input = File.OpenRead(path);
            OfficeDocumentReadResult result = useStream
                ? reader.ReadDocument(input, Path.GetFileName(path), options)
                : reader.ReadDocument(path, options);

            Assert.Contains(result.Chunks, chunk => chunk.Kind == ReaderInputKind.Excel &&
                chunk.Location.Sheet == "Geometry" && chunk.Text?.Contains("Width 16", StringComparison.Ordinal) == true);

            var ordinary = Encoding.UTF8.GetBytes("<Workbook xmlns=\"urn:example\"><item>ordinary XML</item></Workbook>");
            OfficeDocumentReadResult generic = reader.ReadDocument(ordinary, "ordinary.xml", options);
            Assert.Contains(generic.Chunks, chunk => chunk.Kind == ReaderInputKind.Xml &&
                chunk.Text?.Contains("ordinary XML", StringComparison.Ordinal) == true);
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false, ReaderDetectionMode.ExtensionOnly)]
    [InlineData(true, ReaderDetectionMode.ExtensionOnly)]
    [InlineData(false, ReaderDetectionMode.ContentWhenUnknown)]
    [InlineData(true, ReaderDetectionMode.ContentWhenUnknown)]
    [InlineData(false, ReaderDetectionMode.PreferContent)]
    [InlineData(true, ReaderDetectionMode.PreferContent)]
    public void NativeExcelProjectionPreservesEveryDetectionMode(bool useStream, ReaderDetectionMode mode) {
        byte[] bytes;
        using (var document = global::OfficeIMO.Excel.ExcelDocument.Create()) {
            var sheet = document.AddSheet("Native");
            sheet.Cell(1, 1).SetValue("Label");
            sheet.Cell(2, 1).SetValue("native workbook marker");
            bytes = document.ToBytes();
        }
        string path = Path.Combine(Path.GetTempPath(), "officeimo-native-excel-" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.WriteAllBytes(path, bytes);
        try {
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddExcelHandler().AddXmlHandler().Build();
            var options = new ReaderOptions { DetectionMode = mode };
            using var input = File.OpenRead(path);
            OfficeDocumentReadResult result = useStream
                ? reader.ReadDocument(input, Path.GetFileName(path), options)
                : reader.ReadDocument(path, options);
            Assert.Contains(result.Chunks, chunk => chunk.Kind == ReaderInputKind.Excel &&
                chunk.Location.Sheet == "Native" && chunk.Text?.Contains("native workbook marker", StringComparison.Ordinal) == true);
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("sheet.xml")]
    [InlineData("sheet.xlsx")]
    public void PreferContentUsesTheSameSmallByteProbeForSpreadsheetXmlAdmission(string sourceName) {
        string xml = "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\">" +
            new string(' ', 8192) +
            "<Worksheet xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" ss:Name=\"Sheet1\"><Table>" +
            "<Row><Cell><Data ss:Type=\"String\">bounded probe marker</Data></Cell></Row>" +
            "</Table></Worksheet></Workbook>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddExcelHandler().AddXmlHandler().Build();
        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(xml), sourceName,
            new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent, DetectionMaxProbeBytes = 128 });
        Assert.Contains(result.Chunks, chunk => chunk.Kind == ReaderInputKind.Excel &&
            chunk.Text?.Contains("bounded probe marker", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void WordMlWarningsReachReaderWhenTheFirstBlockIsATable() {
        const string xml = """
            <w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml">
              <w:body><w:tbl><w:tblPr><w:unmapped/></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>WordML table</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body>
            </w:wordDocument>
            """;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().AddXmlHandler().Build();
        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(xml), "word.xml",
            new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent });
        Assert.Contains(result.Chunks, chunk => chunk.Kind == ReaderInputKind.Word &&
            chunk.Text?.Contains("WordML table", StringComparison.Ordinal) == true &&
            chunk.Warnings?.Any(warning => warning.Contains("WORDML2003_OMITTED", StringComparison.Ordinal)) == true);
    }

    [Fact]
    public void DocumentReaderXml_ReadXml_ParsesXmlIntoStructuredChunks() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-xml-" + Guid.NewGuid().ToString("N") + ".xml");
        try {
            File.WriteAllText(path,
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<catalog><book id=\"b1\"><title>One</title></book><book id=\"b2\"><title>Two</title></book></catalog>");

            var chunks = XmlReaderAdapter.Read(
                path,
                xmlOptions: new XmlReadOptions {
                    ChunkRows = 2,
                    IncludeMarkdown = true
                }).ToList();

            Assert.NotEmpty(chunks);
            Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Xml, c.Kind));
            Assert.Contains(chunks, c => (c.Text ?? string.Empty).Contains("catalog[1]/book[1]/@id", StringComparison.Ordinal));
            Assert.Contains(chunks, c => c.Tables != null && c.Tables.Count > 0 && c.Tables[0].Columns.Contains("Type", StringComparer.Ordinal));
            Assert.All(chunks.SelectMany(c => c.Tables ?? Array.Empty<ReaderTable>()), table => {
                Assert.Equal(3, table.ColumnProfiles.Count);
                Assert.Equal("Path", table.ColumnProfiles[0].Name);
                Assert.Equal("Type", table.ColumnProfiles[1].Name);
                Assert.Equal("Value", table.ColumnProfiles[2].Name);
            });
            Assert.All(chunks, c => {
                Assert.False(string.IsNullOrWhiteSpace(c.SourceId));
                Assert.False(string.IsNullOrWhiteSpace(c.SourceHash));
                Assert.False(string.IsNullOrWhiteSpace(c.ChunkHash));
                Assert.True(c.TokenEstimate.HasValue && c.TokenEstimate.Value >= 1);
                Assert.True(c.SourceLengthBytes.HasValue && c.SourceLengthBytes.Value > 0);
                Assert.True(c.SourceLastWriteUtc.HasValue);
            });
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReaderXml_ReadXmlStream_NonSeekable_EnforcesMaxInputBytes() {
        var xml =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<catalog><book id=\"b1\"><title>One</title></book><book id=\"b2\"><title>Two</title></book></catalog>";
        using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(xml));

        var ex = Assert.Throws<IOException>(() => XmlReaderAdapter.Read(
            stream,
            sourceName: "books.xml",
            readerOptions: new ReaderOptions { MaxInputBytes = 16 },
            xmlOptions: new XmlReadOptions {
                ChunkRows = 2,
                IncludeMarkdown = true
            }).ToList());

        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderXml_ReadXml_EmitsWarningForMalformedXml() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-bad-xml-" + Guid.NewGuid().ToString("N") + ".xml");
        try {
            File.WriteAllText(path, "<root><broken></root>");

            var chunks = XmlReaderAdapter.Read(path).ToList();

            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Xml &&
                (c.Warnings?.Any(w => w.Contains("XML parse error", StringComparison.OrdinalIgnoreCase)) ?? false));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReaderXml_ReadXml_PreservesCDataAsElementText() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-cdata-xml-" + Guid.NewGuid().ToString("N") + ".xml");
        try {
            File.WriteAllText(path,
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<catalog><book><summary><![CDATA[Line <b>One</b> & more]]></summary></book></catalog>");

            var chunk = Assert.Single(XmlReaderAdapter.Read(
                path,
                xmlOptions: new XmlReadOptions {
                    ChunkRows = 10,
                    IncludeMarkdown = true
                }));

            Assert.Contains("catalog[1]/book[1]/summary[1] | element | Line <b>One</b> & more", chunk.Text ?? string.Empty, StringComparison.Ordinal);
            Assert.Contains("Line <b>One</b> & more", chunk.Markdown ?? string.Empty, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReaderXml_ReadXml_PreservesQualifiedNamesInPaths() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-ns-xml-" + Guid.NewGuid().ToString("N") + ".xml");
        try {
            File.WriteAllText(path,
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<root xmlns:a=\"urn:alpha\" xmlns:b=\"urn:beta\">" +
                "<a:item code=\"A1\">One</a:item>" +
                "<b:item b:code=\"B1\">Two</b:item>" +
                "</root>");

            var chunk = Assert.Single(XmlReaderAdapter.Read(
                path,
                xmlOptions: new XmlReadOptions {
                    ChunkRows = 20,
                    IncludeMarkdown = true
                }));

            Assert.Contains("root[1]/a:item[1] | element | One", chunk.Text ?? string.Empty, StringComparison.Ordinal);
            Assert.Contains("root[1]/b:item[1] | element | Two", chunk.Text ?? string.Empty, StringComparison.Ordinal);
            Assert.Contains("root[1]/b:item[1]/@b:code | attribute | B1", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReaderXml_ReadXml_IndexesAlternatingSiblingNamesIndependently() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-siblings-" + Guid.NewGuid().ToString("N") + ".xml");
        try {
            File.WriteAllText(path, "<root><a>1</a><b>1</b><a>2</a><b>2</b><a>3</a><c>1</c><d>1</d><e>1</e><e>2</e></root>");

            var chunk = Assert.Single(XmlReaderAdapter.Read(
                path,
                xmlOptions: new XmlReadOptions {
                    ChunkRows = 20,
                    IncludeMarkdown = true
                }));

            var text = chunk.Text ?? string.Empty;
            Assert.Contains("root[1]/a[1] | element | 1", text, StringComparison.Ordinal);
            Assert.Contains("root[1]/a[2] | element | 2", text, StringComparison.Ordinal);
            Assert.Contains("root[1]/a[3] | element | 3", text, StringComparison.Ordinal);
            Assert.Contains("root[1]/b[1] | element | 1", text, StringComparison.Ordinal);
            Assert.Contains("root[1]/b[2] | element | 2", text, StringComparison.Ordinal);
            Assert.Contains("root[1]/e[1] | element | 1", text, StringComparison.Ordinal);
            Assert.Contains("root[1]/e[2] | element | 2", text, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
