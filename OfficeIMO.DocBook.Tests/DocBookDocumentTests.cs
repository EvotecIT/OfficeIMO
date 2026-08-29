namespace OfficeIMO.DocBook.Tests;

public sealed class DocBookDocumentTests {
    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void CreatesEditsValidatesAndReopensBothProfiles(DocBookProfile profile) {
        DocBookDocument document = DocBookDocument.CreateArticle(profile); document.Title = "Guide";
        DocBookNode section = document.AddSection("Start"); section.AddParagraph("Hello");
        section.AddItemizedList().AddListItem("One");
        section.AddProgramListing("dotnet test", "shell");
        section.AddAdmonition(DocBookNodeKind.Note, "Remember");
        section.AddImage("image.png", "Example");
        section.AddTable("Data").Add(DocBookNodeKind.TableGroup).Add(DocBookNodeKind.TableBody)
            .Add(DocBookNodeKind.Row).Add(DocBookNodeKind.Entry, "Value");
        section.AddIndexTerm("install");

        Assert.True(document.Validate().IsValid);
        DocBookDocument reopened = DocBookDocument.Parse(document.ToDocBook());
        Assert.Equal(profile, reopened.Profile);
        Assert.Equal("Guide", reopened.Title);
        Assert.Contains(reopened.Xml.Descendants(), e => e.Name.LocalName == "programlisting");
    }

    [Fact]
    public void ExactSchemaProfileAndValidationScopeAreExposed() {
        DocBookValidationResult four = DocBookDocument.CreateBook(DocBookProfile.DocBook45).Validate();
        Assert.Equal("-//OASIS//DTD DocBook XML V4.5//EN", four.SchemaProfile.DtdPublicId);
        Assert.False(four.IsOfficialSchemaValidated);
        DocBookValidationResult five = DocBookDocument.CreateBook(DocBookProfile.DocBook52).Validate();
        Assert.EndsWith("/rng/docbook.rng", five.SchemaProfile.RelaxNgUri);
        Assert.EndsWith("/sch/docbook.sch", five.SchemaProfile.SchematronUri);
    }

    [Fact]
    public void UnchangedInputIsExactAndEditPreservesExtensions() {
        const string source = "<?xml version=\"1.0\"?><article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:test\" version=\"5.2\"><info><title>T</title><x:meta key=\"v\"/></info><!--keep--><section x:flag=\"yes\"><title>S</title><x:block>payload</x:block></section></article>";
        DocBookDocument document = DocBookDocument.Parse(source);
        Assert.Equal(source, document.ToDocBook());
        document.Title = "Changed";
        string edited = document.ToDocBook();
        Assert.Contains("x:meta", edited); Assert.Contains("x:block", edited); Assert.Contains("x:flag=\"yes\"", edited); Assert.Contains("<!--keep-->", edited);
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "DB010" && d.Severity == DocBookDiagnosticSeverity.Info);
    }

    [Fact]
    public void LimitsAndEntityExpansionAreBounded() {
        Assert.Throws<InvalidDataException>(() => DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para/><para/></article>",
            new DocBookReadOptions { MaxElements = 2 }));
        const string entity = "<!DOCTYPE article [<!ENTITY a \"1234567890\">]><article><para>&a;&a;</para></article>";
        Assert.ThrowsAny<Exception>(() => DocBookDocument.Parse(entity, new DocBookReadOptions { MaxCharactersFromEntities = 10 }));
        DocBookDocument external = DocBookDocument.Parse(
            "<!DOCTYPE article [<!ENTITY external SYSTEM \"file:///etc/passwd\">]><article><para>&external;</para></article>");
        Assert.Equal(string.Empty, external.Root.Text);
        using var canceled = new System.Threading.CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => DocBookDocument.Parse("<article><para>P</para></article>", cancellationToken: canceled.Token));
    }

    [Fact]
    public void SharedModelRoundTripRetainsTypedHierarchy() {
        DocBookDocument document = DocBookDocument.CreateBook(DocBookProfile.DocBook45); document.Title = "T";
        document.AddSection("S").AddParagraph("P");
        var model = document.ToOfficeDocumentModel().Value;
        Assert.Equal(OfficeDocumentFormat.DocBook, model.Format);
        Assert.Contains(model.Structure, n => n.Kind == "section");
        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);
        Assert.False(converted.HasLoss);
        Assert.Equal(DocBookDocumentKind.Book, converted.Value.Kind);
        Assert.Equal(DocBookProfile.DocBook45, converted.Value.Profile);
        Assert.Contains(converted.Value.Xml.Descendants(), e => e.Name.LocalName == "section");
        Assert.True(converted.Value.Validate().IsValid);
    }

    [Fact]
    public void SharedModelRoundTripPreservesMixedLinkOrderAndExtensionStructure() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:test\" version=\"5.2\"><para>See <link xmlns:xl=\"http://www.w3.org/1999/xlink\" xl:href=\"https://example.test/\">site</link> now.</para><x:box mode=\"a\">before<x:item>inside</x:item>after</x:box></article>";
        DocBookDocument document = DocBookDocument.Parse(source);
        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(document.ToOfficeDocumentModel().Value);
        Assert.False(converted.HasLoss);
        XElement paragraph = converted.Value.Xml.Descendants().Single(e => e.Name.LocalName == "para");
        Assert.Equal("See site now.", paragraph.Value);
        Assert.Equal("https://example.test/", paragraph.Elements().Single(e => e.Name.LocalName == "link").Attributes().Single(a => a.Name.LocalName == "href").Value);
        XElement extension = converted.Value.Xml.Descendants().Single(e => e.Name == XName.Get("box", "urn:test"));
        Assert.Equal("beforeinsideafter", extension.Value);
        Assert.Equal("a", (string?)extension.Attribute("mode"));
    }

    [Fact]
    public void SharedModelPreservesInformalTablesAndMixedCodeWhitespace() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup><tbody><row><entry>V</entry></row></tbody></tgroup></informaltable><programlisting>before <emphasis>inside</emphasis> after</programlisting><screen>left <replaceable>value</replaceable> right</screen></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;

        Assert.Equal("informal-table", model.Structure[0].Kind);
        DocBookDocument restored = DocBookDocument.FromOfficeDocumentModel(model).Value;
        Assert.Contains(restored.Xml.Descendants(), element => element.Name.LocalName == "informaltable");
        XElement listing = restored.Xml.Descendants().Single(element => element.Name.LocalName == "programlisting");
        Assert.Equal("before inside after", listing.Value);
        Assert.Equal(new[] { "before ", "inside", " after" }, listing.Nodes().Select(node => node is XText text ? text.Value : ((XElement)node).Value));
        XElement screen = restored.Xml.Descendants().Single(element => element.Name.LocalName == "screen");
        Assert.Equal(new[] { "left ", "value", " right" }, screen.Nodes().Select(node => node is XText text ? text.Value : ((XElement)node).Value));
    }

    [Fact]
    public void SharedConversionReportsDocBookFiveVersionNormalization() {
        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.0\"><para>P</para></article>").ToOfficeDocumentModel();

        Assert.True(converted.HasLoss);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB111");
        Assert.Contains("version=\"5.2\"", DocBookDocument.FromOfficeDocumentModel(converted.Value).Value.ToDocBook());
    }

    [Fact]
    public void SharedConversionReportsNativeOnlyComments() {
        DocBookDocument document = DocBookDocument.Parse("<!DOCTYPE article [<!ENTITY marker \"x\">]><article custom=\"x\">root-text<!--native--><para>P</para></article>");
        DocBookConversionResult<OfficeDocumentModel> result = document.ToOfficeDocumentModel();
        Assert.True(result.HasLoss);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DB105");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DB106");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DB107");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DB110");
    }

    [Fact]
    public void SharedConversionKeepsNamespacedLookalikeExtensionsAndDocBook4Ulink() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:extension\" version=\"5.2\"><x:section x:flag=\"yes\">Extension</x:section></article>";
        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(source).ToOfficeDocumentModel();
        OfficeDocumentModelNode extension = Assert.Single(converted.Value.Structure);
        Assert.Equal("extension:{urn:extension}section", extension.Kind);

        DocBookDocument restored = DocBookDocument.FromOfficeDocumentModel(converted.Value).Value;
        Assert.Contains("urn:extension", restored.ToDocBook());

        DocBookDocument docBook4 = DocBookDocument.CreateArticle(DocBookProfile.DocBook45);
        docBook4.AddParagraph("See ").AddLink("site", "https://example.test");
        DocBookDocument roundTripped = DocBookDocument.FromOfficeDocumentModel(
            docBook4.ToOfficeDocumentModel().Value,
            DocBookDocumentKind.Article,
            DocBookProfile.DocBook45).Value;
        Assert.Contains("<ulink", roundTripped.ToDocBook());
        Assert.Contains("url=\"https://example.test\"", roundTripped.ToDocBook());
    }

    [Fact]
    public async System.Threading.Tasks.Task StreamLoadAndWriteHonorCallerOwnership() {
        byte[] bytes = Encoding.UTF8.GetBytes("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para>P</para></article>");
        using var input = new MemoryStream(bytes); input.Position = 3;
        DocBookDocument document = await DocBookDocument.LoadAsync(input);
        Assert.Equal(3, input.Position);
        using var output = new MemoryStream(new byte[512], writable: true);
        await document.WriteAsync(output);
        Assert.Equal(0, output.Position);
        Assert.Equal(DocBookProfile.DocBook52, DocBookDocument.Load(output).Profile);
    }

    [Fact]
    public void LoadedBytesHonorDeclaredEncodingAndParsedTextWritesConsistentUtf8() {
        const string latinSource = "<?xml version=\"1.0\" encoding=\"iso-8859-1\"?><article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para>Caf\u00e9</para></article>";
        byte[] latinBytes = Encoding.GetEncoding("iso-8859-1").GetBytes(latinSource);
        DocBookDocument loaded = DocBookDocument.Load(new MemoryStream(latinBytes));
        Assert.Equal(latinSource, loaded.ToDocBook());
        using var exact = new MemoryStream();
        loaded.Write(exact);
        Assert.Equal(latinBytes, exact.ToArray());

        const string utf16Declaration = "<?xml version=\"1.0\" encoding=\"utf-16\"?><article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"/>";
        DocBookDocument parsed = DocBookDocument.Parse(utf16Declaration);
        using var output = new MemoryStream();
        parsed.Write(output);
        string serialized = Encoding.UTF8.GetString(output.ToArray());
        Assert.Contains("encoding=\"utf-8\"", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(DocBookProfile.DocBook52, DocBookDocument.Load(new MemoryStream(output.ToArray())).Profile);
    }

    [Fact]
    public async System.Threading.Tasks.Task PathSaveAndAsyncLoadReopenTheCommittedArtifact() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-docbook-" + Guid.NewGuid().ToString("N") + ".docbook");
        try {
            DocBookDocument document = DocBookDocument.CreateBook(); document.Title = "Path";
            await document.SaveAsync(path);
            using var stream = File.OpenRead(path);
            Assert.Equal("Path", (await DocBookDocument.LoadAsync(stream)).Title);
            Assert.Equal("Path", DocBookDocument.Load(path).Title);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void AdvancedXmlMutationCannotBeHiddenByUnchangedSourceFastPath() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para>Before</para></article>";
        DocBookDocument document = DocBookDocument.Parse(source);
        document.Xml.Descendants().Single(e => e.Name.LocalName == "para").Value = "After";
        Assert.True(document.IsModified);
        Assert.Contains(">After</", document.ToDocBook());
    }

    [Fact]
    public void DocBookFourLinkUsesUlinkAndExactDtdIsRequiredForCleanProfile() {
        DocBookDocument document = DocBookDocument.CreateArticle(DocBookProfile.DocBook45);
        document.AddParagraph("See ").AddLink("site", "https://example.test/");
        Assert.Contains(document.Xml.Descendants(), e => e.Name.LocalName == "ulink" && (string?)e.Attribute("url") == "https://example.test/");
        Assert.DoesNotContain(document.Validate().Diagnostics, d => d.Code == "DB005");

        DocBookDocument undeclared = DocBookDocument.Parse("<article><para>P</para></article>");
        Assert.Contains(undeclared.Validate().Diagnostics, d => d.Code == "DB005");
    }

    [Fact]
    public void EditingDirectDocBookFourTitleDoesNotCreateDuplicateMetadataTitle() {
        DocBookDocument document = DocBookDocument.Parse("<article><title>Before</title><para>P</para></article>");

        document.Title = "After";

        XElement title = Assert.Single(document.Xml.Root!.Elements(), element => element.Name.LocalName == "title");
        Assert.Equal("After", title.Value);
        Assert.DoesNotContain(document.Xml.Root!.Elements(), element => element.Name.LocalName == "articleinfo");

        DocBookDocument restored = DocBookDocument.FromOfficeDocumentModel(document.ToOfficeDocumentModel().Value).Value;
        Assert.Single(restored.Xml.Root!.Elements(), element => element.Name.LocalName == "title");
        Assert.DoesNotContain(restored.Xml.Root!.Elements(), element => element.Name.LocalName == "articleinfo");
    }
}
