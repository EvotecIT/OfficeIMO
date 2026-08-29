using OfficeIMO.DocBook;
using OfficeIMO.Opml;
using OfficeIMO.Reader.All;
using OfficeIMO.Reader.DocBook;
using OfficeIMO.Reader.Opml;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderOpmlDocBookModularTests {
    [Fact]
    public void OpmlAdapterEmitsNestedOutlineChunksAndRegistersExtension() {
        OpmlDocument document = OpmlDocument.Create();
        document.AddOutline("Root").AddChild("Child");

        ReaderChunk[] chunks = OpmlReaderAdapter.Read(document, "tree.opml").ToArray();
        Assert.Equal(new[] { "Root", "Child" }, chunks.Select(chunk => chunk.Text));
        Assert.Equal("Root > Child", chunks[1].Location.HeadingPath);
        Assert.All(chunks, chunk => Assert.Equal(ReaderInputKind.Opml, chunk.Kind));

        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddOpmlHandler().Build();
        Assert.Equal(ReaderInputKind.Opml, reader.DetectKind("tree.opml"));
    }

    [Fact]
    public void DocBookAdapterEmitsCommonStructureChunksAndRegistersDedicatedExtensions() {
        DocBookDocument document = DocBookDocument.CreateArticle();
        document.AddSection("Start").AddParagraph("Body");

        ReaderChunk[] chunks = DocBookReaderAdapter.Read(document, "guide.docbook").ToArray();
        Assert.Contains(chunks, chunk => chunk.Text == "Start" && chunk.Location.SourceBlockKind == "section");
        Assert.Contains(chunks, chunk => chunk.Text == "Body" && chunk.Location.SourceBlockKind == "paragraph");
        Assert.All(chunks, chunk => Assert.Equal(ReaderInputKind.DocBook, chunk.Kind));

        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddDocBookHandler().Build();
        Assert.Equal(ReaderInputKind.DocBook, reader.DetectKind("guide.dbk"));
        Assert.Equal(ReaderInputKind.DocBook, reader.DetectKind("guide.docbook"));
        Assert.Equal(ReaderInputKind.Xml, reader.DetectKind("guide.xml"));
    }

    [Fact]
    public void DocBookAdapterDoesNotDuplicateAggregateTitleOrExtensionText() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:test\" version=\"5.2\"><info><title>Title</title></info><x:box>Extension</x:box></article>";
        ReaderChunk[] chunks = DocBookReaderAdapter.Read(DocBookDocument.Parse(source)).ToArray();

        Assert.Equal(new[] { "Title", "Extension" }, chunks.Select(chunk => chunk.Text));

        const string compound = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:test\" version=\"5.2\"><x:box>before<x:item>inside</x:item>after</x:box></article>";
        ReaderChunk[] compoundChunks = DocBookReaderAdapter.Read(DocBookDocument.Parse(compound)).ToArray();
        Assert.Equal(new[] { "before", "inside", "after" }, compoundChunks.Select(chunk => chunk.Text));
    }

    [Fact]
    public void AllPresetIncludesBothBoundedAdapters() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();
        Assert.Contains(reader.GetCapabilities(), capability => capability.Id == "officeimo.reader.opml");
        Assert.Contains(reader.GetCapabilities(), capability => capability.Id == "officeimo.reader.docbook");

        byte[] opmlBytes = Encoding.UTF8.GetBytes("<opml version=\"2.0\"><head/><body><outline text=\"A\"/></body></opml>");
        OfficeDocumentReadResult opml = reader.ReadDocument(opmlBytes, "a.opml");
        Assert.Equal(ReaderInputKind.Opml, opml.Kind);
        Assert.Equal("A", Assert.Single(opml.Chunks).Text);
        Assert.Equal(ReaderInputKind.Opml, reader.Detect(opmlBytes, "renamed.bin", new ReaderDetectionOptions {
            Mode = ReaderDetectionMode.PreferContent
        }).Kind);

        byte[] docBookBytes = Encoding.UTF8.GetBytes("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para>P</para></article>");
        OfficeDocumentReadResult docbook = reader.ReadDocument(docBookBytes, "a.docbook");
        Assert.Equal(ReaderInputKind.DocBook, docbook.Kind);
        Assert.Equal("P", Assert.Single(docbook.Chunks).Text);
        OfficeDocumentReadResult renamed = reader.ReadDocument(docBookBytes, "a.xml", new ReaderOptions {
            DetectionMode = ReaderDetectionMode.PreferContent
        });
        Assert.Equal(ReaderInputKind.DocBook, renamed.Kind);
        Assert.Equal("P", Assert.Single(renamed.Chunks).Text);

        byte[] prefixedDocBook = Encoding.UTF8.GetBytes("<?xml version=\"1.0\"?><db:article xmlns:db=\"http://docbook.org/ns/docbook\" version=\"5.2\"><db:para>Q</db:para></db:article>");
        Assert.Equal(ReaderInputKind.DocBook, reader.Detect(prefixedDocBook, "renamed.bin", new ReaderDetectionOptions {
            Mode = ReaderDetectionMode.PreferContent
        }).Kind);

        byte[] ordinaryXml = Encoding.UTF8.GetBytes("<root><value>&lt;opml version=\"2.0\"&gt;</value></root>");
        Assert.NotEqual(ReaderInputKind.Opml, reader.Detect(ordinaryXml, "renamed.bin", new ReaderDetectionOptions {
            Mode = ReaderDetectionMode.PreferContent
        }).Kind);

        string[] lookalikes = {
            "<OpMl version=\"2.0\"><head/><body/></OpMl>",
            "<opml xmlns=\"urn:not-opml\" version=\"2.0\"><head/><body/></opml>",
            "<Article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"/>",
            "<article xmlns=\"urn:not-docbook\" data-format=\"http://docbook.org/ns/docbook\"/>",
            "<book data-format=\"http://docbook.org/ns/docbook\"/>",
            "<!-- -//OASIS//DTD DocBook XML V4.5//EN http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd --><article/>",
            "<!DOCTYPE article [<!-- -//OASIS//DTD DocBook XML V4.5//EN http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd -->]><article/>",
            "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN EXTRA\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article/>",
            "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd?other\"><article/>",
            "<!DOCTYPE book PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article/>"
        };
        foreach (string lookalike in lookalikes) {
            ReaderDetectionResult detection = reader.Detect(Encoding.UTF8.GetBytes(lookalike), "renamed.bin", new ReaderDetectionOptions {
                Mode = ReaderDetectionMode.PreferContent
            });
            Assert.NotEqual(ReaderInputKind.Opml, detection.Kind);
            Assert.NotEqual(ReaderInputKind.DocBook, detection.Kind);
        }

        byte[] docBook45 = Encoding.UTF8.GetBytes("<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><para>P</para></article>");
        Assert.Equal(ReaderInputKind.DocBook, reader.Detect(docBook45, "renamed.bin", new ReaderDetectionOptions {
            Mode = ReaderDetectionMode.PreferContent
        }).Kind);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ContentDetectionRecognizesUtf32XmlInBothByteOrders(bool bigEndian) {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();
        var encoding = new UTF32Encoding(bigEndian, byteOrderMark: true, throwOnInvalidCharacters: true);

        byte[] opml = encoding.GetPreamble().Concat(encoding.GetBytes("<opml version=\"2.0\"><head/><body/></opml>")).ToArray();
        byte[] docBook = encoding.GetPreamble().Concat(encoding.GetBytes("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"/>")).ToArray();
        var options = new ReaderDetectionOptions { Mode = ReaderDetectionMode.PreferContent };

        Assert.Equal(ReaderInputKind.Opml, reader.Detect(opml, "renamed.bin", options).Kind);
        Assert.Equal(ReaderInputKind.DocBook, reader.Detect(docBook, "renamed.bin", options).Kind);
    }

    [Fact]
    public void ContentDetectionScansTheConfiguredXmlProbeForDelayedRoots() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();
        string prefix = "<!--" + new string('x', 6_000) + "-->";
        var options = new ReaderDetectionOptions { Mode = ReaderDetectionMode.PreferContent, MaxProbeBytes = 8_192 };

        Assert.Equal(ReaderInputKind.Opml, reader.Detect(
            Encoding.UTF8.GetBytes(prefix + "<opml version=\"2.0\"><head/><body/></opml>"), "renamed.bin", options).Kind);
        Assert.Equal(ReaderInputKind.DocBook, reader.Detect(
            Encoding.UTF8.GetBytes(prefix + "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"/>"), "renamed.bin", options).Kind);
    }

    [Theory]
    [InlineData(ReaderInputKind.Opml, OfficeDocumentFormat.Opml)]
    [InlineData(ReaderInputKind.DocBook, OfficeDocumentFormat.DocBook)]
    public void PdfBridgeRetainsNewReaderFormats(ReaderInputKind readerKind, OfficeDocumentFormat expectedFormat) {
        Assert.Equal(expectedFormat, OfficeDocumentReadResultPdfExtensions.MapFormat(readerKind));
    }

    [Fact]
    public void RegistrationsPublishNativeLimitsBeforeSnapshottingStreams() {
        var opmlOptions = new ReaderOpmlOptions { ReadOptions = new OpmlReadOptions { MaxInputBytes = 8 } };
        OfficeDocumentReader opmlReader = new OfficeDocumentReaderBuilder().AddOpmlHandler(opmlOptions).Build();
        Assert.Equal(8, opmlReader.GetHandlerDefaultMaxInputBytes("input.opml"));
        using var opmlStream = new global::OfficeIMO.Tests.NonSeekableReadStream(new byte[9]);
        IOException opmlException = Assert.Throws<IOException>(() => opmlReader.ReadDocument(opmlStream, "input.opml"));
        Assert.Contains("MaxInputBytes", opmlException.Message, StringComparison.Ordinal);

        var docBookOptions = new ReaderDocBookOptions { ReadOptions = new DocBookReadOptions { MaxInputBytes = 12 } };
        OfficeDocumentReader docBookReader = new OfficeDocumentReaderBuilder().AddDocBookHandler(docBookOptions).Build();
        Assert.Equal(12, docBookReader.GetHandlerDefaultMaxInputBytes("input.docbook"));
        using var docBookStream = new global::OfficeIMO.Tests.NonSeekableReadStream(new byte[13]);
        IOException docBookException = Assert.Throws<IOException>(() => docBookReader.ReadDocument(docBookStream, "input.docbook"));
        Assert.Contains("MaxInputBytes", docBookException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ReaderLimitIsAppliedBeforeNativeParsing() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("<opml version=\"2.0\"><head/><body><outline text=\"too long\"/></body></opml>"));
        Assert.Throws<InvalidDataException>(() => OpmlReaderAdapter.Read(stream, "a.opml", new ReaderOptions { MaxInputBytes = 8 }).ToArray());
        using var canceled = new System.Threading.CancellationTokenSource(); canceled.Cancel();
        using var second = new MemoryStream(Encoding.UTF8.GetBytes("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"/>"));
        Assert.Throws<OperationCanceledException>(() => DocBookReaderAdapter.Read(second, "a.docbook", cancellationToken: canceled.Token).ToArray());
    }
}
