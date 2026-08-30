using OfficeIMO.DocBook;
using OfficeIMO.Opml;
using OfficeIMO.Reader.All;
using OfficeIMO.Reader.DocBook;
using OfficeIMO.Reader.Opml;
using System.Threading;
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
    public void OpmlAdapterPublishesEveryDistinctOutlineTargetInChunkMarkdown() {
        OpmlDocument document = OpmlDocument.Create();
        OpmlOutline outline = document.AddOutline("Feed");
        outline.XmlUrl = "https://example.test/feed.xml";
        outline.HtmlUrl = "https://example.test/site (home)";
        outline.Url = "https://example.test/feed.xml";

        ReaderChunk chunk = Assert.Single(OpmlReaderAdapter.Read(document));

        Assert.Equal("Feed", chunk.Text);
        Assert.Contains("- Feed: [https://example.test/feed.xml](https://example.test/feed.xml)", chunk.Markdown, StringComparison.Ordinal);
        Assert.Contains("- Website: [https://example.test/site (home)](https://example.test/site%20\\(home\\))", chunk.Markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("- Link:", chunk.Markdown, StringComparison.Ordinal);
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
    public void DocBookAdapterPublishesBoundedCalsTablesAndConversionWarnings() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup cols=\"2\"><colspec colname=\"c1\"/><colspec colname=\"c2\"/><thead><row><entry namest=\"c1\" nameend=\"c2\">Values</entry></row></thead><tbody><row><entry>A</entry><entry>1</entry></row><row><entry>B</entry><entry>2</entry></row></tbody></tgroup></informaltable></article>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddDocBookHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(
            Encoding.UTF8.GetBytes(source), "table.docbook", new ReaderOptions { MaxTableRows = 1 });
        ReaderTable table = Assert.Single(result.Tables);

        Assert.Equal(new[] { "Values", "Column 2" }, table.Columns);
        Assert.Equal(new[] { "A", "1" }, Assert.Single(table.Rows));
        Assert.Equal(2, table.TotalRowCount);
        Assert.True(table.Truncated);
        Assert.Contains(result.Chunks, chunk => chunk.Tables?.Count == 1);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DB112" &&
            diagnostic.Category == OfficeDocumentDiagnosticCategory.Adapter);
        Assert.All(result.Chunks, chunk => Assert.Null(chunk.Warnings));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DB113" &&
            diagnostic.Category == OfficeDocumentDiagnosticCategory.Limit);
    }

    [Fact]
    public void DocBookAdapterDoesNotDuplicateAggregateTitleOrExtensionText() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:test\" version=\"5.2\"><info><title>Title</title></info><x:box>Extension</x:box></article>";
        ReaderChunk[] chunks = DocBookReaderAdapter.Read(DocBookDocument.Parse(source)).ToArray();

        Assert.Equal(new[] { "Title", "Extension" }, chunks.Select(chunk => chunk.Text));

        const string compound = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:test\" version=\"5.2\"><x:box>before<x:item>inside</x:item>after</x:box></article>";
        ReaderChunk[] compoundChunks = DocBookReaderAdapter.Read(DocBookDocument.Parse(compound)).ToArray();
        Assert.Equal(new[] { "before", "inside", "after" }, compoundChunks.Select(chunk => chunk.Text));

        ReaderChunk[] directContainer = DocBookReaderAdapter.Read(DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><note>Careful</note></article>")).ToArray();
        Assert.Equal("Careful", Assert.Single(directContainer).Text);
    }

    [Fact]
    public void DocBookAdapterPropagatesAdmonitionContextToNestedContent() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><warning><para>Danger</para></warning><note><itemizedlist><listitem><para>Remember</para></listitem></itemizedlist></note></article>";

        ReaderChunk[] chunks = DocBookReaderAdapter.Read(DocBookDocument.Parse(source)).ToArray();

        Assert.Contains(chunks, chunk => chunk.Text == "Danger" && chunk.Location.SourceBlockKind == "warning");
        Assert.Contains(chunks, chunk => chunk.Text == "Remember" && chunk.Location.SourceBlockKind == "note" && chunk.Markdown == "- Remember");
    }

    [Fact]
    public void DocBookAdapterKeepsContinuationParagraphsInsideListItems() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><itemizedlist><listitem><para>First</para><para>Second</para></listitem></itemizedlist><orderedlist startingnumber=\"12\"><listitem><para>Third</para><para>Fourth</para></listitem></orderedlist></article>";

        ReaderChunk[] chunks = DocBookReaderAdapter.Read(DocBookDocument.Parse(source)).ToArray();

        Assert.Contains(chunks, chunk => chunk.Markdown == "- First");
        Assert.Contains(chunks, chunk => chunk.Markdown == "  Second");
        Assert.Contains(chunks, chunk => chunk.Markdown == "12. Third");
        Assert.Contains(chunks, chunk => chunk.Markdown == "    Fourth");
    }

    [Fact]
    public void DocBookAdapterEmitsParentMarkerBeforeLeadingNestedList() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><itemizedlist><listitem><itemizedlist><listitem><para>Nested</para></listitem></itemizedlist><para>After</para></listitem></itemizedlist></article>";

        ReaderChunk[] chunks = DocBookReaderAdapter.Read(DocBookDocument.Parse(source)).ToArray();

        Assert.Equal(new[] { "-", "  - Nested", "  After" }, chunks.Select(chunk => chunk.Markdown));
        Assert.Equal(new[] { string.Empty, "Nested", "After" }, chunks.Select(chunk => chunk.Text));
    }

    [Fact]
    public void DocBookAdapterRendersInlineLinkAndCrossReferenceTargets() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:xl=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><para>See <link xl:href=\"https://example.test/a_(b)\">site</link> and <xref linkend=\"target\"/>.</para></article>";

        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(source));
        OfficeDocumentReadResult result = DocBookReaderAdapter.ReadDocument(stream);

        Assert.Equal("See [site](https://example.test/a_\\(b\\)) and [target](#target).", result.Markdown);
        Assert.Contains(result.Chunks, chunk => chunk.Markdown == "[site](https://example.test/a_\\(b\\))");
        Assert.Contains(result.Chunks, chunk => chunk.Markdown == "[target](#target)");

        ReaderChunk legacyLink = Assert.Single(DocBookReaderAdapter.Read(DocBookDocument.Parse(
            "<article><ulink url=\"https://example.test/legacy\">legacy</ulink></article>")));
        Assert.Equal("[legacy](https://example.test/legacy)", legacyLink.Markdown);

        ReaderChunk whitespaceLink = Assert.Single(DocBookReaderAdapter.Read(DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:xl=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><link xl:href=\"https://example.test/a b&#x9;c&#xA0;d\">space</link></article>")));
        Assert.Equal("[space](https://example.test/a%20b%09c%C2%A0d)", whitespaceLink.Markdown);
    }

    [Fact]
    public void DocBookAdapterScopesInlineProjectionToTheStructuralTitle() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:xl=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><section><title>See <link xl:href=\"https://example.test\">site</link></title><para>Body</para></section></article>";

        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(source));
        OfficeDocumentReadResult result = DocBookReaderAdapter.ReadDocument(stream);
        string markdown = result.Markdown.Replace("\r\n", "\n");

        Assert.Contains("# See [site](https://example.test)", markdown, StringComparison.Ordinal);
        Assert.Contains("\n\nBody", markdown, StringComparison.Ordinal);
        Assert.Contains(result.Chunks, chunk => chunk.Text == "Body" && chunk.Location.SourceBlockKind == "paragraph");
        Assert.DoesNotContain(result.Chunks, chunk => chunk.Text.IndexOf("Body", StringComparison.Ordinal) >= 0 && chunk.Text != "Body");
    }

    [Fact]
    public void DocBookAdapterTraversesCompoundExtensionContentInOrder() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:test\" version=\"5.2\"><x:box>before<x:item>inside</x:item>after</x:box></article>";

        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(source));
        OfficeDocumentReadResult result = DocBookReaderAdapter.ReadDocument(stream);

        Assert.Equal(new[] { "before", "inside", "after" }, result.Chunks.Select(chunk => chunk.Text));
        Assert.DoesNotContain(result.Chunks, chunk => chunk.Text == "beforeinsideafter");
    }

    [Fact]
    public void DocBookAdapterUsesAFenceThatCannotBeClosedByListingContent() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><programlisting>before\n```\n# still code\nafter</programlisting></article>";

        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(source));
        OfficeDocumentReadResult result = DocBookReaderAdapter.ReadDocument(stream);
        string markdown = result.Markdown.Replace("\r\n", "\n");

        Assert.StartsWith("~~~\nbefore\n```\n# still code\nafter\n~~~", markdown, StringComparison.Ordinal);
        Assert.Equal("code", Assert.Single(result.Chunks).Location.SourceBlockKind);
    }

    [Fact]
    public void DocBookAdapterFencesScreenContentAsPreformattedText() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><screen># prompt\n```\nstill screen</screen></article>";

        ReaderChunk chunk = Assert.Single(DocBookReaderAdapter.Read(DocBookDocument.Parse(source)));

        Assert.Equal("screen", chunk.Location.SourceBlockKind);
        Assert.Equal("~~~\n# prompt\n```\nstill screen\n~~~", chunk.Markdown.Replace("\r\n", "\n"));
    }

    [Fact]
    public void DocBookAdapterKeepsInlineTargetsInsideProgramListingFences() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:xl=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><programlisting>before <link xl:href=\"https://example.test\">target</link> after</programlisting></article>";

        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(source));
        OfficeDocumentReadResult result = DocBookReaderAdapter.ReadDocument(stream);
        ReaderChunk chunk = Assert.Single(result.Chunks);

        Assert.Equal("before target after", chunk.Text);
        Assert.Equal("```\nbefore target after\n```", chunk.Markdown.Replace("\r\n", "\n"));
        Assert.DoesNotContain("[target]", chunk.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void DocBookAdapterObservesCancellationDuringSemanticProjection() {
        DocBookDocument document = DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para>Body</para></article>");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            DocBookReaderAdapter.Read(document, cancellationToken: cancellation.Token).ToArray());
    }

    [Fact]
    public void DocBookAdapterExcludesIndexTermsFromMarkdown() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para>Body<indexterm><primary>topic</primary></indexterm></para></article>";

        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(source));
        OfficeDocumentReadResult result = DocBookReaderAdapter.ReadDocument(stream);

        Assert.Equal("Body", result.Markdown);
    }

    [Fact]
    public void DocBookAdapterAppliesAndClassifiesTheAggregateTextBudget() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para><link xmlns:xl=\"http://www.w3.org/1999/xlink\" xl:href=\"https://example.test\"><emphasis>1234567890</emphasis></link></para></article>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(source));

        OfficeDocumentReadResult result = DocBookReaderAdapter.ReadDocument(stream, docBookOptions: new ReaderDocBookOptions {
            ConversionOptions = new DocBookConversionOptions { MaxTotalTextCharacters = 12 }
        });

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DB123" &&
            diagnostic.Category == OfficeDocumentDiagnosticCategory.Limit);
    }

    [Fact]
    public void DocBookAdapterPreservesItemizedOrderedAndNestedListMarkers() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><itemizedlist><listitem><para>Alpha</para><itemizedlist><listitem><para>Nested</para></listitem></itemizedlist></listitem><listitem><para>Beta</para></listitem></itemizedlist><orderedlist startingnumber=\"3\"><listitem><para>Third</para></listitem><listitem><para>Fourth</para></listitem></orderedlist><orderedlist startingnumber=\"100\"><listitem><para>Hundred</para><itemizedlist><listitem><para>Nested hundred</para></listitem></itemizedlist></listitem></orderedlist></article>";

        ReaderChunk[] chunks = DocBookReaderAdapter.Read(DocBookDocument.Parse(source)).ToArray();

        Assert.Contains(chunks, chunk => chunk.Markdown == "- Alpha" && chunk.Location.SourceBlockKind == "list-item");
        Assert.Contains(chunks, chunk => chunk.Markdown == "  - Nested" && chunk.Location.SourceBlockKind == "list-item");
        Assert.Contains(chunks, chunk => chunk.Markdown == "- Beta" && chunk.Location.SourceBlockKind == "list-item");
        Assert.Contains(chunks, chunk => chunk.Markdown == "3. Third" && chunk.Location.SourceBlockKind == "list-item");
        Assert.Contains(chunks, chunk => chunk.Markdown == "4. Fourth" && chunk.Location.SourceBlockKind == "list-item");
        Assert.Contains(chunks, chunk => chunk.Markdown == "100. Hundred" && chunk.Location.SourceBlockKind == "list-item");
        Assert.Contains(chunks, chunk => chunk.Markdown == "     - Nested hundred" && chunk.Location.SourceBlockKind == "list-item");

        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(source));
        string markdown = DocBookReaderAdapter.ReadDocument(stream).Markdown.Replace("\r\n", "\n");
        Assert.Contains("- Alpha", markdown, StringComparison.Ordinal);
        Assert.Contains("  - Nested", markdown, StringComparison.Ordinal);
        Assert.Contains("3. Third", markdown, StringComparison.Ordinal);
        Assert.Contains("     - Nested hundred", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void OpmlAndDocBookAdaptersKeepSurrogatePairsIntactWhenSplitting() {
        OpmlDocument opml = OpmlDocument.Create();
        opml.AddOutline("A😀B");
        ReaderChunk[] opmlChunks = OpmlReaderAdapter.Read(opml, readerOptions: new ReaderOptions { MaxChars = 2 }).ToArray();
        Assert.Equal(new[] { "A", "😀", "B" }, opmlChunks.Select(chunk => chunk.Text));
        Assert.Equal("A😀B", string.Concat(opmlChunks.Select(chunk => chunk.Text)));
        Assert.False(opmlChunks[0].ContinuesPreviousChunk);
        Assert.All(opmlChunks.Skip(1), chunk => Assert.True(chunk.ContinuesPreviousChunk));
        Assert.Equal(new[] { "# A", "😀", "B" }, opmlChunks.Select(chunk => chunk.Markdown));

        using var opmlStream = new MemoryStream(Encoding.UTF8.GetBytes(opml.ToOpml()));
        OfficeDocumentReadResult opmlResult = OpmlReaderAdapter.ReadDocument(opmlStream, readerOptions: new ReaderOptions { MaxChars = 2 });
        Assert.Equal("# A😀B", opmlResult.Markdown);

        DocBookDocument docBook = DocBookDocument.CreateArticle();
        docBook.AddParagraph("A😀B");
        ReaderChunk[] docBookChunks = DocBookReaderAdapter.Read(docBook, readerOptions: new ReaderOptions { MaxChars = 2 }).ToArray();
        Assert.Equal(new[] { "A", "😀", "B" }, docBookChunks.Select(chunk => chunk.Text));
        Assert.Equal("A😀B", string.Concat(docBookChunks.Select(chunk => chunk.Text)));
        Assert.False(docBookChunks[0].ContinuesPreviousChunk);
        Assert.All(docBookChunks.Skip(1), chunk => Assert.True(chunk.ContinuesPreviousChunk));

        using var docBookStream = new MemoryStream(Encoding.UTF8.GetBytes(docBook.ToDocBook()));
        OfficeDocumentReadResult docBookResult = DocBookReaderAdapter.ReadDocument(docBookStream, readerOptions: new ReaderOptions { MaxChars = 2 });
        Assert.Equal("A😀B", docBookResult.Markdown);

        DocBookDocument codeBook = DocBookDocument.CreateArticle();
        codeBook.Root.Add(DocBookNodeKind.ProgramListing, "A😀B");
        using var codeStream = new MemoryStream(Encoding.UTF8.GetBytes(codeBook.ToDocBook()));
        OfficeDocumentReadResult codeResult = DocBookReaderAdapter.ReadDocument(codeStream, readerOptions: new ReaderOptions { MaxChars = 2 });
        Assert.Equal("```\nA😀B\n```", codeResult.Markdown);
    }

    [Fact]
    public void OpmlAdapterAttachesDocumentWarningsToOnlyOneChunk() {
        const string source = "<opml version=\"9.0\"><head/><body><outline text=\"One\"/><outline text=\"Two\"/></body></opml>";

        ReaderChunk[] chunks = OpmlReaderAdapter.Read(OpmlDocument.Parse(source)).ToArray();

        Assert.Equal(2, chunks.Length);
        Assert.Single(chunks, chunk => chunk.Warnings?.Count > 0);
    }

    [Fact]
    public void OpmlAdapterBoundsCumulativeHeadingPaths() {
        OpmlDocument document = OpmlDocument.Create();
        document.AddOutline(new string('a', 2_000)).AddChild(new string('b', 2_000));

        ReaderChunk[] chunks = OpmlReaderAdapter.Read(document).ToArray();

        Assert.All(chunks, chunk => Assert.True(chunk.Location.HeadingPath!.Length <= 1_024));
    }

    [Fact]
    public void DocBookAdapterAttachesDocumentWarningsToOnlyOneChunk() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><ulink url=\"https://example.test\">Link</ulink><para>One</para><para>Two</para></article>";

        ReaderChunk[] chunks = DocBookReaderAdapter.Read(DocBookDocument.Parse(source)).ToArray();

        Assert.True(chunks.Length > 1);
        Assert.Single(chunks, chunk => chunk.Warnings?.Count > 0);
    }

    [Fact]
    public void DocBookAdapterRetainsBoundedSectionHeadingPaths() {
        DocBookDocument document = DocBookDocument.CreateArticle();
        document.AddSection(new string('a', 2_000)).AddSection(new string('b', 2_000));

        ReaderChunk[] chunks = DocBookReaderAdapter.Read(document).ToArray();

        Assert.All(chunks, chunk => Assert.True((chunk.Location.HeadingPath?.Length ?? 0) <= 1_024));
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
        byte[] referencedNamespace = Encoding.UTF8.GetBytes("<article xmlns=\"http://docbook.org/ns/docboo&#x6B;\" version=\"5.2\"/>");
        Assert.Equal(ReaderInputKind.DocBook, reader.Detect(referencedNamespace, "renamed.bin", new ReaderDetectionOptions {
            Mode = ReaderDetectionMode.PreferContent
        }).Kind);
        byte[] entityBackedAttribute = Encoding.UTF8.GetBytes("<!DOCTYPE article [<!ENTITY role \"guide\">]><article xmlns=\"http://docbook.org/ns/docbook\" role=\"&role;\"><para>P</para></article>");
        Assert.Equal(ReaderInputKind.DocBook, reader.Detect(entityBackedAttribute, "renamed.bin", new ReaderDetectionOptions {
            Mode = ReaderDetectionMode.PreferContent
        }).Kind);
        OfficeDocumentReadResult entityBackedResult = reader.ReadDocument(entityBackedAttribute, "renamed.bin", new ReaderOptions {
            DetectionMode = ReaderDetectionMode.PreferContent
        });
        Assert.Equal(ReaderInputKind.DocBook, entityBackedResult.Kind);
        Assert.Equal("P", Assert.Single(entityBackedResult.Chunks).Text);

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
        byte[] docBook45WithSubsetMarkup = Encoding.UTF8.GetBytes(
            "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\" [<!-- ]> --><!ENTITY role \"guide\"><?officeimo ]>?>]><article role=\"&role;\"><para>P</para></article>");
        Assert.Equal(ReaderInputKind.DocBook, reader.Detect(docBook45WithSubsetMarkup, "renamed.bin", new ReaderDetectionOptions {
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
    public void ContentDetectionRecognizesBomlessUtf16AndUtf32XmlByteOrders() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();
        Encoding[] encodings = {
            new UnicodeEncoding(false, false),
            new UnicodeEncoding(true, false),
            new UTF32Encoding(false, false),
            new UTF32Encoding(true, false)
        };

        foreach (Encoding encoding in encodings) {
            byte[] opml = encoding.GetBytes("<opml version=\"2.0\"><head/><body><outline text=\"A\"/></body></opml>");
            byte[] docBook = encoding.GetBytes("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para>P</para></article>");
            var options = new ReaderDetectionOptions { Mode = ReaderDetectionMode.PreferContent };

            Assert.Equal(ReaderInputKind.Opml, reader.Detect(opml, "renamed.bin", options).Kind);
            Assert.Equal(ReaderInputKind.DocBook, reader.Detect(docBook, "renamed.bin", options).Kind);
        }
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

    [Fact]
    public void OpmlAdapterAppliesCallerConversionBounds() {
        OpmlDocument document = OpmlDocument.Create();
        document.AddOutline("One");
        document.AddOutline("Two");
        var options = new ReaderOpmlOptions {
            ConversionOptions = new OpmlConversionOptions { MaxStructureNodes = 1 }
        };

        Assert.Throws<InvalidDataException>(() => OpmlReaderAdapter.Read(document, opmlOptions: options).ToArray());
    }

    [Fact]
    public void OpmlRichResultPublishesMetadataLinksAndDocumentDiagnosticsOnce() {
        const string source = "<opml version=\"9.0\"><head><title>Feeds</title><ownerName>Jane Doe</ownerName></head><body><outline text=\"Feed\" type=\"rss\" xmlUrl=\"https://example.test/feed.xml\" htmlUrl=\"https://example.test/\"><extension/></outline><outline text=\"Missing\" type=\"link\"/></body></opml>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddOpmlHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(source), "feeds.opml");

        Assert.Equal("Feeds", result.Source.Title);
        Assert.Equal("Jane Doe", result.Source.Author);
        Assert.Contains(result.Metadata, entry => entry.Name == "version" && entry.Value == "9.0");
        Assert.Contains(result.Links, link => link.Kind == "subscription" && link.Uri == "https://example.test/feed.xml");
        Assert.Contains(result.Links, link => link.Kind == "html" && link.Uri == "https://example.test/");
        Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "OPML001");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML001" &&
            diagnostic.Category == OfficeDocumentDiagnosticCategory.Parsing);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML012" &&
            diagnostic.Category == OfficeDocumentDiagnosticCategory.Content);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML200" &&
            diagnostic.Category == OfficeDocumentDiagnosticCategory.Adapter);
        Assert.Contains(result.Metadata, entry => entry.Category == "reader.summary" &&
            entry.Name == "ChunkCount" && entry.Value == "2");
        Assert.All(result.Chunks, chunk => Assert.Null(chunk.Warnings));
    }

    [Fact]
    public void DocBookRichResultPublishesAuthorLinksTablesAndDocumentDiagnosticsOnce() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:xl=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><info><title>Guide</title><author>Jane Doe</author></info><para><link xl:href=\"https://example.test\">Site</link></para><informaltable><tgroup cols=\"2\"><thead><row><entry namest=\"c1\" nameend=\"c2\">Header</entry></row></thead><tbody><row><entry>A</entry><entry>B</entry></row></tbody><colspec colname=\"c1\"/><colspec colname=\"c2\"/></tgroup></informaltable></article>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddDocBookHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(source), "guide.docbook");

        Assert.Equal("Guide", result.Source.Title);
        Assert.Equal("Jane Doe", result.Source.Author);
        Assert.Contains(result.Metadata, entry => entry.Name == "author" && entry.Value == "Jane Doe");
        Assert.Contains(result.Links, link => link.Uri == "https://example.test" && link.Text == "Site");
        Assert.Single(result.Tables);
        Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "DB112");
        Assert.All(result.Chunks, chunk => Assert.Null(chunk.Warnings));
    }

    [Fact]
    public void DocBookRichResultKeepsDocumentTitleDepthTableContextAndImageReferences() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><info><title>Guide</title></info><section><title>Details</title><table><title>Values</title><tgroup><tbody><row><entry>A</entry></row></tbody></tgroup></table><mediaobject><imageobject><imagedata fileref=\"assets/figure.png\"/></imageobject><caption><para>Chart</para></caption></mediaobject></section></article>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddDocBookHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(source), "guide.docbook");
        ReaderTable table = Assert.Single(result.Tables);
        OfficeDocumentAsset asset = Assert.Single(result.Assets);

        Assert.StartsWith("# Guide", result.Markdown, StringComparison.Ordinal);
        Assert.Contains("\n\n# Details", result.Markdown.Replace("\r\n", "\n"), StringComparison.Ordinal);
        Assert.DoesNotContain("## Guide", result.Markdown, StringComparison.Ordinal);
        Assert.Equal("Details / Values", table.Location!.HeadingPath);
        Assert.Equal("assets/figure.png", asset.SourceObjectId);
        Assert.Equal("figure.png", asset.FileName);
        Assert.Equal("Chart", asset.Title);
        Assert.Equal("Details", asset.Location.HeadingPath);
        Assert.Equal("image", asset.Location.SourceBlockKind);
        Assert.Null(asset.PayloadBytes);
        Assert.Contains(result.Metadata, entry => entry.Category == "reader.summary" &&
            entry.Name == "AssetCount" && entry.Value == "1");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "ocr-needed");
    }

    [Fact]
    public void DocBookRichResultExcludesNestedEntryRowsAndExtensionAlternativeText() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:extension\" version=\"5.2\"><informaltable><tgroup cols=\"1\"><tbody><row><entry>Outer<entrytbl cols=\"1\"><tbody><row><entry>Nested</entry></row></tbody></entrytbl></entry></row></tbody></tgroup></informaltable><mediaobject><imageobject><imagedata fileref=\"figure.png\"/></imageobject><x:textobject><x:phrase>Internal data</x:phrase></x:textobject></mediaobject></article>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddDocBookHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(source), "nested.docbook");

        ReaderTable table = Assert.Single(result.Tables);
        Assert.Equal(1, table.TotalRowCount);
        Assert.Single(table.Rows);
        Assert.Null(Assert.Single(result.Assets).AltText);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DB121");
    }

    [Fact]
    public void DocBookRichResultDoesNotPromoteSectionAuthorToDocumentSource() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><section><title>Chapter</title><info><author>Jane Doe</author></info></section></article>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddDocBookHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(source), "guide.docbook");

        Assert.Null(result.Source.Author);
        Assert.Contains(result.Metadata, entry => entry.Name == "author" && entry.Value == "Jane Doe");
    }

    [Fact]
    public void DocBookRichResultClassifiesParsingAndContentDiagnostics() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.1\"><section/></article>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddDocBookHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(source), "invalid.docbook");

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DB003" &&
            diagnostic.Category == OfficeDocumentDiagnosticCategory.Parsing);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DB011" &&
            diagnostic.Category == OfficeDocumentDiagnosticCategory.Content);
    }

    [Fact]
    public void RichResultsRetainDiagnosticsWhenChunkWarningsAreDisabled() {
        OfficeDocumentReader opmlReader = new OfficeDocumentReaderBuilder()
            .AddOpmlHandler(new ReaderOpmlOptions { IncludeDiagnostics = false }).Build();
        OfficeDocumentReadResult opml = opmlReader.ReadDocument(Encoding.UTF8.GetBytes(
            "<opml version=\"9.0\"><head/><body><outline text=\"Item\"/></body></opml>"), "invalid.opml");
        Assert.Contains(opml.Diagnostics, diagnostic => diagnostic.Code == "OPML001");
        Assert.All(opml.Chunks, chunk => Assert.Null(chunk.Warnings));

        OfficeDocumentReader docBookReader = new OfficeDocumentReaderBuilder()
            .AddDocBookHandler(new ReaderDocBookOptions { IncludeDiagnostics = false }).Build();
        OfficeDocumentReadResult docBook = docBookReader.ReadDocument(Encoding.UTF8.GetBytes(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><ulink url=\"https://example.test\">Link</ulink></article>"), "invalid.docbook");
        Assert.Contains(docBook.Diagnostics, diagnostic => diagnostic.Code == "DB014");
        Assert.All(docBook.Chunks, chunk => Assert.Null(chunk.Warnings));
    }

    [Fact]
    public void RichResultsRetainDiagnosticsWhenNoChunksAreProduced() {
        OfficeDocumentReader opmlReader = new OfficeDocumentReaderBuilder().AddOpmlHandler().Build();
        OfficeDocumentReadResult opml = opmlReader.ReadDocument(
            Encoding.UTF8.GetBytes("<opml version=\"2.0\"><head/><body><extension/></body></opml>"), "empty.opml");
        Assert.Empty(opml.Chunks);
        Assert.Contains(opml.Diagnostics, diagnostic => diagnostic.Code == "OPML203");

        OfficeDocumentReader docBookReader = new OfficeDocumentReaderBuilder().AddDocBookHandler().Build();
        OfficeDocumentReadResult docBook = docBookReader.ReadDocument(
            Encoding.UTF8.GetBytes("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><ulink url=\"https://example.test\"/></article>"), "empty.docbook");
        Assert.Empty(docBook.Chunks);
        Assert.Contains(docBook.Diagnostics, diagnostic => diagnostic.Code == "DB014");
    }

    [Fact]
    public void ChunkOnlyReadersEmitDiagnosticChunksWhenNoContentExists() {
        ReaderChunk opml = Assert.Single(OpmlReaderAdapter.Read(OpmlDocument.Parse(
            "<opml version=\"9.0\"><head/><body/></opml>")));
        Assert.Equal("diagnostic", opml.Location.SourceBlockKind);
        Assert.Contains(opml.Warnings!, warning => warning.StartsWith("OPML001:", StringComparison.Ordinal));

        ReaderChunk docBook = Assert.Single(DocBookReaderAdapter.Read(DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.1\"/>")));
        Assert.Equal("diagnostic", docBook.Location.SourceBlockKind);
        Assert.Contains(docBook.Warnings!, warning => warning.StartsWith("DB003:", StringComparison.Ordinal));
    }

    [Fact]
    public void RichResultsPreserveDistinctDiagnosticOccurrences() {
        OfficeDocumentReader opmlReader = new OfficeDocumentReaderBuilder().AddOpmlHandler().Build();
        OfficeDocumentReadResult opml = opmlReader.ReadDocument(Encoding.UTF8.GetBytes(
            "<opml version=\"2.0\"><head/><body><outline text=\"Same\"><extension/></outline><outline text=\"Same\"><extension/></outline></body></opml>"), "duplicates.opml");
        Assert.Equal(2, opml.Diagnostics.Count(diagnostic => diagnostic.Code == "OPML200"));

        OfficeDocumentReader docBookReader = new OfficeDocumentReaderBuilder().AddDocBookHandler().Build();
        OfficeDocumentReadResult docBook = docBookReader.ReadDocument(Encoding.UTF8.GetBytes(
            "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:extension\" version=\"5.2\"><x:box/><x:box/></article>"), "duplicates.docbook");
        Assert.Equal(2, docBook.Diagnostics.Count(diagnostic => diagnostic.Code == "DB100"));
    }

    [Fact]
    public void OpmlRichResultBoundsRepeatedExtensionDiagnostics() {
        string outlines = string.Concat(Enumerable.Range(1, 105)
            .Select(index => $"<outline text=\"{index}\"><extension/></outline>"));
        string source = $"<opml version=\"2.0\"><head/><body>{outlines}</body></opml>";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddOpmlHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(source), "extensions.opml");

        Assert.Equal(101, result.Diagnostics.Count(diagnostic => diagnostic.Code == "OPML200"));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML200" &&
            diagnostic.Message.StartsWith("5 additional", StringComparison.Ordinal));
    }
}
