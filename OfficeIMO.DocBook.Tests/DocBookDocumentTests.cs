using System.Collections.Generic;

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
    public void ValidationReportsAMissingRootInsteadOfThrowing() {
        DocBookDocument document = DocBookDocument.CreateArticle();
        document.Xml.Root!.Remove();

        DocBookValidationResult validation = document.Validate();

        Assert.False(validation.IsValid);
        Assert.Contains(validation.Diagnostics, diagnostic => diagnostic.Code == "DB001" &&
            diagnostic.Severity == DocBookDiagnosticSeverity.Error);
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
    public void LimitsAndEntityPolicyAreBounded() {
        Assert.Throws<InvalidDataException>(() => DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para/><para/></article>",
            new DocBookReadOptions { MaxElements = 2 }));
        const string entity = "<!DOCTYPE article [<!ENTITY a \"1234567890\">]><article><para>&a;&a;</para></article>";
        Assert.ThrowsAny<Exception>(() => DocBookDocument.Parse(entity, new DocBookReadOptions { MaxCharactersFromEntities = 10 }));
        InvalidDataException external = Assert.Throws<InvalidDataException>(() => DocBookDocument.Parse(
            "<!DOCTYPE article [<!ENTITY external SYSTEM \"file:///etc/passwd\">]><article><para>&external;</para></article>"));
        Assert.Contains("external and parameter entity declarations", external.Message, StringComparison.Ordinal);
        Assert.ThrowsAny<Exception>(() => DocBookDocument.Parse(
            "<!DOCTYPE article [<!ENTITY % external SYSTEM \"file:///etc/passwd\">]><article/>"));
        using var canceled = new System.Threading.CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => DocBookDocument.Parse("<article><para>P</para></article>", cancellationToken: canceled.Token));
    }

    [Fact]
    public void ValidationRejectsListItemsOutsideSupportedListParents() {
        DocBookDocument invalid = DocBookDocument.CreateArticle();
        invalid.Root.AddListItem("Root item");
        invalid.AddSection("Section").AddListItem("Section item");

        Assert.Equal(2, invalid.Validate().Diagnostics.Count(diagnostic => diagnostic.Code == "DB015"));

        DocBookDocument variableList = DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><variablelist><varlistentry><term>Name</term><listitem><para>Value</para></listitem></varlistentry></variablelist></article>");
        Assert.DoesNotContain(variableList.Validate().Diagnostics, diagnostic => diagnostic.Code == "DB015");
    }

    [Fact]
    public void RepeatedDiagnosticsAreCappedPerCodeAndSummarized() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:extension\" version=\"5.2\"><x:a/><x:b/><x:c/><x:d/><x:e/></article>";
        DocBookDocument document = DocBookDocument.Parse(source);

        DocBookValidationResult validation = document.Validate(new DocBookValidationOptions {
            MaxDetailedDiagnosticsPerCode = 2
        });
        DocBookConversionResult<OfficeDocumentModel> conversion = document.ToOfficeDocumentModel(options: new DocBookConversionOptions {
            MaxDetailedDiagnosticsPerCode = 2
        });
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Structure = Enumerable.Range(1, 5)
                .Select(index => new OfficeDocumentModelNode { Kind = "unsupported-" + index, Text = "Value" }).ToArray()
        };
        DocBookConversionResult<DocBookDocument> reverse = DocBookDocument.FromOfficeDocumentModel(
            model, options: new DocBookConversionOptions { MaxDetailedDiagnosticsPerCode = 2 });

        Assert.Equal(3, validation.Diagnostics.Count(diagnostic => diagnostic.Code == "DB010"));
        Assert.Contains(validation.Diagnostics, diagnostic => diagnostic.Code == "DB010" &&
            diagnostic.Message.StartsWith("3 additional", StringComparison.Ordinal));
        Assert.Equal(3, conversion.Diagnostics.Count(diagnostic => diagnostic.Code == "DB100"));
        Assert.Contains(conversion.Diagnostics, diagnostic => diagnostic.Code == "DB100" &&
            diagnostic.Message.StartsWith("3 additional", StringComparison.Ordinal));
        Assert.Equal(3, reverse.Diagnostics.Count(diagnostic => diagnostic.Code == "DB101"));
        Assert.Contains(reverse.Diagnostics, diagnostic => diagnostic.Code == "DB101" &&
            diagnostic.Message.StartsWith("3 additional", StringComparison.Ordinal));
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
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup><thead><row><entry>Name</entry><entry>Value</entry></row></thead><tbody><row><entry>A</entry><entry>1</entry></row></tbody></tgroup></informaltable><programlisting>before <emphasis>inside</emphasis> after</programlisting><screen>left <replaceable>value</replaceable> right</screen></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;

        Assert.Equal("informal-table", model.Structure[0].Kind);
        OfficeDocumentModelTable table = Assert.Single(model.Tables);
        Assert.Equal("informaltable", table.Kind);
        Assert.Equal(new[] { "Name", "Value" }, table.Columns);
        Assert.Equal(new[] { "A", "1" }, Assert.Single(table.Rows));
        DocBookDocument restored = DocBookDocument.FromOfficeDocumentModel(model).Value;
        Assert.Single(restored.Xml.Descendants(), element => element.Name.LocalName == "informaltable");
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

    [Theory]
    [InlineData("<article><para>P</para></article>")]
    [InlineData("<!DOCTYPE book PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><para>P</para></article>")]
    public void SharedConversionReportsDocBookFourDoctypeNormalization(string source) {
        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(source).ToOfficeDocumentModel();

        Assert.True(converted.HasLoss);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB107");
        XDocumentType restored = Assert.IsType<XDocumentType>(
            DocBookDocument.FromOfficeDocumentModel(converted.Value).Value.Xml.DocumentType);
        Assert.Equal("article", restored.Name);
        Assert.Equal(DocBookSchemaProfiles.DocBook45.DtdPublicId, restored.PublicId);
        Assert.Equal(DocBookSchemaProfiles.DocBook45.DtdSystemId, restored.SystemId);
    }

    [Fact]
    public void SharedTableProjectionAlignsCalsSpansAndConsumesAllHeaderRows() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup cols=\"3\"><colspec colname=\"c1\"/><colspec colname=\"c2\"/><colspec colname=\"c3\"/><thead><row><entry namest=\"c1\" nameend=\"c2\">Group</entry><entry>C</entry></row><row><entry>A</entry><entry>B</entry><entry>C2</entry></row></thead><tbody><row><entry namest=\"c1\" nameend=\"c2\">Wide</entry><entry>Last</entry></row><row><entry morerows=\"1\">Tall</entry><entry>R1B</entry><entry>R1C</entry></row><row><entry>R2B</entry><entry>R2C</entry></row></tbody></tgroup></informaltable></article>";

        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(source).ToOfficeDocumentModel();
        OfficeDocumentModelTable table = Assert.Single(converted.Value.Tables);

        Assert.True(converted.HasLoss);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB112");
        Assert.Equal(new[] { "Group / A", "B", "C / C2" }, table.Columns);
        Assert.Equal(new[] { "Wide", "", "Last" }, table.Rows[0]);
        Assert.Equal(new[] { "Tall", "R1B", "R1C" }, table.Rows[1]);
        Assert.Equal(new[] { "", "R2B", "R2C" }, table.Rows[2]);
        Assert.Equal(3, table.TotalRowCount);
    }

    [Fact]
    public void SharedTableProjectionReportsFooterFlattening() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup cols=\"1\"><tbody><row><entry>Body</entry></row></tbody><tfoot><row><entry>Total</entry></row></tfoot></tgroup></informaltable></article>";

        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(source).ToOfficeDocumentModel();
        OfficeDocumentModelTable table = Assert.Single(converted.Value.Tables);

        Assert.Equal(2, table.TotalRowCount);
        Assert.Equal(new[] { "Body", "Total" }, table.Rows.Select(row => row.Single()).ToArray());
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB119");
    }

    [Fact]
    public void SharedTableProjectionReportsMultipleGroupFlattening() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup cols=\"1\"><tbody><row><entry>A</entry></row></tbody></tgroup><tgroup cols=\"2\"><tbody><row><entry>B</entry><entry>C</entry></row></tbody></tgroup></informaltable></article>";

        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(source).ToOfficeDocumentModel();
        OfficeDocumentModelTable table = Assert.Single(converted.Value.Tables);

        Assert.Equal(2, table.TotalRowCount);
        Assert.Equal(new[] { "A", "" }, table.Rows[0]);
        Assert.Equal(new[] { "B", "C" }, table.Rows[1]);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB112");
    }

    [Fact]
    public void SharedTableProjectionExcludesNestedEntryTableRows() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup cols=\"1\"><tbody><row><entry>Outer<entrytbl cols=\"1\"><tbody><row><entry>Nested</entry></row></tbody></entrytbl></entry></row></tbody></tgroup></informaltable></article>";

        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(source).ToOfficeDocumentModel();
        OfficeDocumentModelTable table = Assert.Single(converted.Value.Tables);

        Assert.Equal(1, table.TotalRowCount);
        Assert.Equal("OuterNested", Assert.Single(Assert.Single(table.Rows)));
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB121");
    }

    [Fact]
    public void SharedTableProjectionPlacesImplicitEntriesAfterExplicitColumns() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup cols=\"3\"><colspec colname=\"c1\"/><colspec colname=\"c2\"/><colspec colname=\"c3\"/><tbody><row><entry colname=\"c2\">B</entry><entry>C</entry></row></tbody></tgroup></informaltable></article>";

        OfficeDocumentModelTable table = Assert.Single(DocBookDocument.Parse(source).ToOfficeDocumentModel().Value.Tables);

        Assert.Equal(new[] { "", "B", "C" }, Assert.Single(table.Rows));
    }

    [Fact]
    public void SharedTableProjectionRetainsEnclosingSectionPath() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><section><title>Outer</title><section><title>Inner</title><table><title>Values</title><tgroup><tbody><row><entry>A</entry></row></tbody></tgroup></table><informaltable><tgroup><tbody><row><entry>B</entry></row></tbody></tgroup></informaltable></section></section></article>";

        OfficeDocumentModelTable[] tables = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value.Tables.ToArray();

        Assert.Equal("Outer / Inner / Values", tables[0].Location!.HeadingPath);
        Assert.Equal("Outer / Inner", tables[1].Location!.HeadingPath);
    }

    [Fact]
    public void SharedModelPublishesMetadataOnlyImageReferences() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><section><title>Figures</title><mediaobject><imageobject><imagedata fileref=\"assets/figure.png?version=2\"/></imageobject><textobject><phrase>Chart alternative</phrase></textobject><caption><para>Chart title</para></caption></mediaobject></section></article>";

        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel("guide.docbook").Value;
        OfficeDocumentModelAsset asset = Assert.Single(model.Assets);

        Assert.Equal("image", asset.Kind);
        Assert.Equal("figure.png", asset.FileName);
        Assert.Equal(".png", asset.Extension);
        Assert.Equal("image/png", asset.MediaType);
        Assert.Equal("assets/figure.png?version=2", asset.SourceObjectId);
        Assert.Equal("Chart alternative", asset.AltText);
        Assert.Equal("Chart title", asset.Title);
        Assert.Equal("Figures", asset.Location.HeadingPath);
        Assert.Equal("guide.docbook", asset.Location.Path);
        Assert.Null(asset.PayloadBytes);
        Assert.Contains("docbook.media-references", model.CapabilitiesUsed);
        Assert.Single(DocBookDocument.FromOfficeDocumentModel(model).Value.Xml.Descendants(),
            element => element.Name.LocalName == "imagedata" &&
                       (string?)element.Attribute("fileref") == "assets/figure.png?version=2");
    }

    [Fact]
    public void SharedModelDoesNotUseExtensionTextAsImageAlternativeText() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:extension\" version=\"5.2\"><mediaobject><imageobject><imagedata fileref=\"figure.png\"/></imageobject><x:textobject><x:phrase>Internal data</x:phrase></x:textobject></mediaobject></article>";

        OfficeDocumentModelAsset asset = Assert.Single(DocBookDocument.Parse(source).ToOfficeDocumentModel().Value.Assets);

        Assert.Null(asset.AltText);
        Assert.Null(asset.Title);
    }

    [Fact]
    public void SharedTableProjectionReportsExactTotalBeyondProjectionCapacity() {
        string rows = string.Concat(Enumerable.Range(1, 10)
            .Select(index => $"<row><entry>Row {index}</entry></row>"));
        string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup><tbody>" +
            rows + "</tbody></tgroup></informaltable></article>";

        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(source)
            .ToOfficeDocumentModel(options: new DocBookConversionOptions { MaxTableRows = 1 });
        OfficeDocumentModelTable table = Assert.Single(converted.Value.Tables);

        Assert.Equal(10, table.TotalRowCount);
        Assert.Single(table.Rows);
        Assert.True(table.Truncated);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB113");
    }

    [Fact]
    public void SharedTableProjectionReservesSeparateHeaderAndBodyCapacity() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup><thead><row><entry morerows=\"3\">H1</entry></row><row><entry>H2</entry></row><row><entry>H3</entry></row><row><entry>H4</entry></row></thead><tbody><row><entry>Body 1</entry></row><row><entry>Body 2</entry></row></tbody></tgroup></informaltable></article>";

        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(source)
            .ToOfficeDocumentModel(options: new DocBookConversionOptions { MaxTableRows = 1 });
        OfficeDocumentModelTable table = Assert.Single(converted.Value.Tables);

        Assert.Equal(new[] { "H1" }, table.Columns);
        Assert.Equal(new[] { "Body 1" }, Assert.Single(table.Rows));
        Assert.Equal(2, table.TotalRowCount);
        Assert.True(table.Truncated);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB113");
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
    public void ProfileConversionTranslatesExternalLinkElementAndAttribute() {
        const string docBookFive = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:xl=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><para><link xl:href=\"https://example.test/five\">Five</link></para></article>";
        DocBookDocument asFour = DocBookDocument.FromOfficeDocumentModel(
            DocBookDocument.Parse(docBookFive).ToOfficeDocumentModel().Value,
            profile: DocBookProfile.DocBook45).Value;
        XElement fourLink = asFour.Xml.Descendants().Single(element => element.Name.LocalName == "ulink");
        Assert.Equal("https://example.test/five", (string?)fourLink.Attribute("url"));
        Assert.Null(fourLink.Attribute(XName.Get("href", "http://www.w3.org/1999/xlink")));

        const string docBookFour = "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><para><ulink url=\"https://example.test/four\">Four</ulink></para></article>";
        DocBookDocument asFive = DocBookDocument.FromOfficeDocumentModel(
            DocBookDocument.Parse(docBookFour).ToOfficeDocumentModel().Value,
            profile: DocBookProfile.DocBook52).Value;
        XElement fiveLink = asFive.Xml.Descendants().Single(element => element.Name.LocalName == "link");
        Assert.Equal("https://example.test/four", (string?)fiveLink.Attribute(XName.Get("href", "http://www.w3.org/1999/xlink")));
        Assert.Null(fiveLink.Attribute("url"));
    }

    [Fact]
    public void PrimaryTitleLookupIgnoresNamespacedExtensionTitles() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:test\" version=\"5.2\"><section><x:title>Extension</x:title><para>Body</para></section></article>";

        OfficeDocumentModelNode section = Assert.Single(DocBookDocument.Parse(source).ToOfficeDocumentModel().Value.Structure);

        Assert.Equal("section", section.Kind);
        Assert.Equal(string.Empty, section.Text);
        Assert.Contains(section.Children, child => child.Kind == "extension:{urn:test}title" && child.Text == "Extension");
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

        DocBookDocument wrongRoot = DocBookDocument.Parse(
            "<!DOCTYPE book PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><para>P</para></article>");
        Assert.Contains(wrongRoot.Validate().Diagnostics, d => d.Code == "DB005");
    }

    [Fact]
    public void InformalTableDoesNotRequireATitle() {
        DocBookDocument document = DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup><tbody><row><entry>V</entry></row></tbody></tgroup></informaltable></article>");

        Assert.DoesNotContain(document.Validate().Diagnostics, diagnostic => diagnostic.Code == "DB011");
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

    [Fact]
    public void SharedTableProjectionBoundsHostileCalsGeometry() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup cols=\"2147483647\"><colspec colname=\"far\" colnum=\"2147483647\"/><tbody><row><entry colname=\"far\">Value</entry></row></tbody></tgroup></informaltable></article>";

        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(source)
            .ToOfficeDocumentModel(options: new DocBookConversionOptions { MaxTableColumns = 4, MaxTableRows = 10 });
        OfficeDocumentModelTable table = Assert.Single(converted.Value.Tables);

        Assert.True(converted.HasLoss);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB113");
        Assert.True(table.Columns.Count <= 4);
        Assert.True(table.Truncated);
    }

    [Fact]
    public void SharedTableProjectionBoundsTotalCellSlots() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup cols=\"4\"><tbody><row><entry>A</entry></row><row><entry>B</entry></row><row><entry>C</entry></row></tbody></tgroup></informaltable></article>";

        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(source)
            .ToOfficeDocumentModel(options: new DocBookConversionOptions { MaxTableColumns = 4, MaxTableRows = 10, MaxTableCells = 8 });
        OfficeDocumentModelTable table = Assert.Single(converted.Value.Tables);

        Assert.Equal(4, table.Columns.Count);
        Assert.Equal(2, table.Rows.Count);
        Assert.All(table.Rows, row => Assert.Equal(4, row.Count));
        Assert.Equal(3, table.TotalRowCount);
        Assert.True(table.Truncated);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB113");
        Assert.Throws<ArgumentOutOfRangeException>(() => DocBookDocument.Parse(source)
            .ToOfficeDocumentModel(options: new DocBookConversionOptions { MaxTableCells = 0 }));
    }

    [Fact]
    public void SharedProjectionBoundsAggregateNestedTextMaterialization() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para><link xmlns:xl=\"http://www.w3.org/1999/xlink\" xl:href=\"https://example.test\"><emphasis>1234567890</emphasis></link></para></article>";

        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(source)
            .ToOfficeDocumentModel(options: new DocBookConversionOptions { MaxTotalTextCharacters = 12 });
        OfficeDocumentModelNode root = Assert.Single(converted.Value.Structure);

        Assert.True(Flatten(root).Sum(node => (long)node.Text.Length) <= 12);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB123");
        Assert.Throws<ArgumentOutOfRangeException>(() => DocBookDocument.Parse(source)
            .ToOfficeDocumentModel(options: new DocBookConversionOptions { MaxTotalTextCharacters = 0 }));

        static IEnumerable<OfficeDocumentModelNode> Flatten(OfficeDocumentModelNode node) {
            yield return node;
            foreach (OfficeDocumentModelNode child in node.Children) {
                foreach (OfficeDocumentModelNode descendant in Flatten(child)) yield return descendant;
            }
        }
    }

    [Fact]
    public void SharedProjectionUsesStructuredPersonNameWithoutAffiliationText() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><info><author><personname><firstname>Jane</firstname><surname>Doe</surname></personname><affiliation><orgname>Acme</orgname></affiliation></author></info></article>";

        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;

        Assert.Equal("Jane Doe", model.Source.Author);
        Assert.Contains(model.Metadata, entry => entry.Name == "author" && entry.Value == "Jane Doe");
        Assert.Contains(Flatten(model.Structure), node => node.Text == "Acme");

        static IEnumerable<OfficeDocumentModelNode> Flatten(IEnumerable<OfficeDocumentModelNode> nodes) {
            foreach (OfficeDocumentModelNode node in nodes) {
                yield return node;
                foreach (OfficeDocumentModelNode descendant in Flatten(node.Children)) yield return descendant;
            }
        }
    }

    [Theory]
    [InlineData("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><ulink url=\"https://example.test\">X</ulink></article>", "DB014")]
    [InlineData("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><section><sectioninfo/></section></article>", "DB014")]
    [InlineData("<article><info><title>T</title></info></article>", "DB014")]
    [InlineData("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><entry>X</entry></article>", "DB015")]
    [InlineData("<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:extension\" version=\"5.2\"><x:row><entry>X</entry></x:row></article>", "DB015")]
    [InlineData("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para><link url=\"https://example.test\">X</link></para></article>", "DB016")]
    [InlineData("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para><link href=\"https://example.test\">X</link></para></article>", "DB016")]
    [InlineData("<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><para><link url=\"https://example.test\">X</link></para></article>", "DB016")]
    [InlineData("<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><para><link href=\"https://example.test\">X</link></para></article>", "DB016")]
    public void BoundedValidationRejectsWrongProfileNamesParentsAndLinkTargets(string source, string code) {
        DocBookValidationResult validation = DocBookDocument.Parse(source).Validate();

        Assert.False(validation.IsValid);
        Assert.Contains(validation.Diagnostics, diagnostic => diagnostic.Code == code && diagnostic.Severity == DocBookDiagnosticSeverity.Error);
    }

    [Fact]
    public void SharedConversionReportsUnsupportedIdentityAndNativeAliases() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Metadata = new[] {
                new OfficeDocumentModelMetadataEntry { Category = "docbook", Name = "kind", Value = "chapter" },
                new OfficeDocumentModelMetadataEntry { Category = "docbook", Name = "profile", Value = "6.0" }
            },
            Structure = new[] { new OfficeDocumentModelNode { Kind = "paragraph", Text = "Body" } }
        };
        DocBookConversionResult<DocBookDocument> restored = DocBookDocument.FromOfficeDocumentModel(model);
        Assert.True(restored.HasLoss);
        Assert.Equal(2, restored.Diagnostics.Count(diagnostic => diagnostic.Code == "DB114"));

        DocBookConversionResult<OfficeDocumentModel> aliases = DocBookDocument.Parse(
            "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><sect1><title>S</title><simpara>P</simpara></sect1></article>")
            .ToOfficeDocumentModel();
        Assert.True(aliases.HasLoss);
        Assert.Contains(aliases.Diagnostics, diagnostic => diagnostic.Code == "DB115");

        DocBookConversionResult<OfficeDocumentModel> docBookFiveAlias = DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><ulink url=\"https://example.test\">Site</ulink></article>")
            .ToOfficeDocumentModel();
        Assert.Contains(docBookFiveAlias.Diagnostics, diagnostic => diagnostic.Code == "DB115");

        DocBookConversionResult<OfficeDocumentModel> docBookFourAlias = DocBookDocument.Parse(
            "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><info><title>T</title></info></article>")
            .ToOfficeDocumentModel();
        Assert.Contains(docBookFourAlias.Diagnostics, diagnostic => diagnostic.Code == "DB115");
    }

    [Fact]
    public void SharedConversionUsesSourceTitleWhenMetadataContainsOnlyAuthor() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Source = new OfficeDocumentModelSource { Title = "Guide" },
            Structure = new[] {
                new OfficeDocumentModelNode {
                    Kind = "metadata",
                    Children = new[] { new OfficeDocumentModelNode { Kind = "author", Text = "Jane Doe" } }
                }
            }
        };

        DocBookDocument converted = DocBookDocument.FromOfficeDocumentModel(model).Value;

        Assert.Equal("Guide", converted.Title);
        Assert.Equal("Jane Doe", converted.Xml.Descendants().Single(element => element.Name.LocalName == "author").Value);
        Assert.Single(converted.Xml.Root!.Elements(), element => element.Name.LocalName == "info");
    }

    [Fact]
    public void SharedConversionDoesNotDuplicateTitlesNestedInRootMetadata() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Source = new OfficeDocumentModelSource { Title = "Source title" },
            Structure = new[] {
                new OfficeDocumentModelNode {
                    Kind = "metadata",
                    Children = new[] {
                        new OfficeDocumentModelNode {
                            Kind = "extension:{urn:test}titles",
                            Children = new[] { new OfficeDocumentModelNode { Kind = "title", Text = "Nested title" } }
                        }
                    }
                }
            }
        };

        DocBookDocument converted = DocBookDocument.FromOfficeDocumentModel(model).Value;

        Assert.Equal("Nested title", Assert.Single(converted.Xml.Descendants(), element => element.Name.LocalName == "title").Value);
    }

    [Fact]
    public void SharedConversionEmitsFlatTablesAndSourceAuthorWithoutRecursiveStructure() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Source = new OfficeDocumentModelSource { Author = "Jane Doe" },
            Tables = new[] {
                new OfficeDocumentModelTable {
                    Title = "Values",
                    Columns = new[] { "Name", "Value" },
                    Rows = new[] { new[] { "A", "1" } },
                    TotalRowCount = 1
                }
            },
            Assets = new[] {
                new OfficeDocumentModelAsset {
                    Id = "figure",
                    Kind = "image",
                    FileName = "figure.png",
                    SourceObjectId = "assets/figure.png",
                    Title = "Chart"
                }
            },
            Links = new[] {
                new OfficeDocumentModelLink {
                    Id = "site",
                    Kind = "link",
                    Text = "Site",
                    Uri = "https://example.test/"
                },
                new OfficeDocumentModelLink {
                    Id = "target",
                    Kind = "cross-reference",
                    Text = "Target",
                    DestinationName = "target-id"
                }
            }
        };

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);
        XElement info = Assert.Single(converted.Value.Xml.Root!.Elements(), element => element.Name.LocalName == "info");
        XElement table = Assert.Single(converted.Value.Xml.Root!.Elements(), element => element.Name.LocalName == "table");
        XElement image = Assert.Single(converted.Value.Xml.Descendants(), element => element.Name.LocalName == "imagedata");
        XElement[] links = converted.Value.Xml.Descendants().Where(element => element.Name.LocalName == "link").ToArray();

        Assert.Equal("Jane Doe", info.Descendants().Single(element => element.Name.LocalName == "author").Value);
        Assert.Equal(new[] { "Name", "Value", "A", "1" },
            table.Descendants().Where(element => element.Name.LocalName == "entry").Select(element => element.Value));
        Assert.Equal("assets/figure.png", image.Attribute("fileref")!.Value);
        Assert.Contains(links, link => (string?)link.Attribute(XName.Get("href", "http://www.w3.org/1999/xlink")) == "https://example.test/");
        Assert.Contains(links, link => (string?)link.Attribute("linkend") == "target-id");
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB103");
        OfficeDocumentModel projected = converted.Value.ToOfficeDocumentModel().Value;
        Assert.Contains(projected.Links, link => link.Uri == "https://example.test/" && link.Text == "Site");
        Assert.Contains(projected.Links, link => link.DestinationName == "target-id" && link.Text == "Target");
    }

    [Fact]
    public void SharedConversionAppendsSupplementaryChannelsAlongsideStructure() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Structure = new[] { new OfficeDocumentModelNode { Id = "structure", Kind = "paragraph", Text = "Structured" } },
            Blocks = new[] { new OfficeDocumentModelBlock { Id = "supplemental-block", Kind = "paragraph", Text = "Supplemental" } },
            Tables = new[] {
                new OfficeDocumentModelTable {
                    Title = "Values", Columns = new[] { "Name" }, Rows = new[] { new[] { "A" } }, TotalRowCount = 1
                }
            },
            Assets = new[] {
                new OfficeDocumentModelAsset { Id = "supplemental-image", Kind = "image", SourceObjectId = "image.png" }
            },
            Links = new[] {
                new OfficeDocumentModelLink { Id = "supplemental-link", Kind = "link", Text = "Site", Uri = "https://example.test/" }
            }
        };

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.Contains(converted.Value.Xml.Descendants(), element => element.Name.LocalName == "para" && element.Value == "Supplemental");
        Assert.Single(converted.Value.Xml.Descendants(), element => element.Name.LocalName == "table");
        Assert.Single(converted.Value.Xml.Descendants(), element => element.Name.LocalName == "imagedata");
        Assert.Single(converted.Value.Xml.Descendants(), element => element.Name.LocalName == "link");
        Assert.Equal(4, converted.Diagnostics.Count(diagnostic => diagnostic.Code == "DB122"));
    }

    [Fact]
    public void SharedRoundTripDoesNotDuplicateDerivedSupplementaryChannels() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para>See <link xmlns:xlink=\"http://www.w3.org/1999/xlink\" xlink:href=\"https://example.test/\">Site</link></para><informaltable><tgroup cols=\"1\"><tbody><row><entry>A</entry></row></tbody></tgroup></informaltable><mediaobject><imageobject><imagedata fileref=\"image.png\"/></imageobject></mediaobject></article>";

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(
            DocBookDocument.Parse(source).ToOfficeDocumentModel().Value);

        Assert.DoesNotContain(converted.Diagnostics, diagnostic => diagnostic.Code == "DB122");
        Assert.Single(converted.Value.Xml.Descendants(), element => element.Name.LocalName == "link");
        Assert.Single(converted.Value.Xml.Descendants(), element => element.Name.LocalName == "informaltable");
        Assert.Single(converted.Value.Xml.Descendants(), element => element.Name.LocalName == "imagedata");
    }

    [Fact]
    public void SharedConversionUsesDocBookFourUlinkForFlatExternalLinks() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Metadata = new[] {
                new OfficeDocumentModelMetadataEntry { Category = "docbook", Name = "profile", Value = "4.5" }
            },
            Links = new[] {
                new OfficeDocumentModelLink { Id = "site", Kind = "link", Text = "Site", Uri = "https://example.test/" }
            }
        };

        DocBookDocument converted = DocBookDocument.FromOfficeDocumentModel(model).Value;
        XElement link = Assert.Single(converted.Xml.Descendants(), element => element.Name.LocalName == "ulink");

        Assert.Equal("https://example.test/", (string?)link.Attribute("url"));
        Assert.Equal("Site", link.Value);
    }

    [Fact]
    public void SharedConversionUsesSourceAuthorWhenStructureHasNoDocumentAuthor() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Source = new OfficeDocumentModelSource { Author = "Jane Doe" },
            Structure = new[] { new OfficeDocumentModelNode { Kind = "paragraph", Text = "Body" } }
        };

        DocBookDocument converted = DocBookDocument.FromOfficeDocumentModel(model).Value;

        Assert.Equal("Jane Doe", converted.Xml.Root!.Elements().Single(element => element.Name.LocalName == "info")
            .Descendants().Single(element => element.Name.LocalName == "author").Value);
    }

    [Theory]
    [InlineData("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><info><authorgroup><author>Jane Doe</author></authorgroup></info></article>")]
    [InlineData("<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><articleinfo><authorgroup><author>Jane Doe</author></authorgroup></articleinfo></article>")]
    public void SharedRoundTripDoesNotDuplicateAuthorsNestedInRootMetadata(string source) {
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;

        DocBookDocument restored = DocBookDocument.FromOfficeDocumentModel(model).Value;

        Assert.Equal("Jane Doe", model.Source.Author);
        Assert.Single(restored.Xml.Descendants(), element => element.Name.LocalName == "author");
    }

    [Fact]
    public void SharedConversionDoesNotTreatSectionTitleAsDocumentTitle() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Source = new OfficeDocumentModelSource { Title = "Guide" },
            Structure = new[] {
                new OfficeDocumentModelNode {
                    Kind = "section",
                    Children = new[] { new OfficeDocumentModelNode { Kind = "title", Text = "Chapter" } }
                }
            }
        };

        DocBookDocument converted = DocBookDocument.FromOfficeDocumentModel(model).Value;

        Assert.Equal("Guide", converted.Title);
        Assert.Equal("Chapter", converted.Xml.Descendants().Single(element =>
            element.Name.LocalName == "section").Elements().Single(element => element.Name.LocalName == "title").Value);
    }

    [Fact]
    public void SharedConversionPreservesTextOnEmptyCommonContainers() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Structure = new[] {
                new OfficeDocumentModelNode { Kind = "section", Text = "Chapter" },
                new OfficeDocumentModelNode { Kind = "note", Text = "Careful" }
            }
        };

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        XElement section = converted.Value.Xml.Descendants().Single(element => element.Name.LocalName == "section");
        Assert.Equal("Chapter", section.Elements().Single(element => element.Name.LocalName == "title").Value);
        XElement note = converted.Value.Xml.Descendants().Single(element => element.Name.LocalName == "note");
        Assert.Equal("Careful", note.Elements().Single(element => element.Name.LocalName == "para").Value);
        Assert.DoesNotContain(converted.Diagnostics, diagnostic => diagnostic.Code == "DB116");
    }

    [Fact]
    public void ProfileConversionRequalifiesUntypedDocBookVocabularyButNotExtensions() {
        const string docBookFive = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:extension\" version=\"5.2\"><info><author><personname><firstname>Jane</firstname><surname>Doe</surname></personname></author></info><indexterm><primary>topic</primary></indexterm><x:box><x:item>value</x:item></x:box></article>";
        DocBookDocument asFour = DocBookDocument.FromOfficeDocumentModel(
            DocBookDocument.Parse(docBookFive).ToOfficeDocumentModel().Value,
            profile: DocBookProfile.DocBook45).Value;
        XElement[] fourVocabulary = asFour.Xml.Descendants()
            .Where(element => new[] { "personname", "firstname", "surname", "primary" }.Contains(element.Name.LocalName)).ToArray();
        Assert.Equal(4, fourVocabulary.Length);
        Assert.All(fourVocabulary,
            element => Assert.Equal(XNamespace.None, element.Name.Namespace));
        Assert.Equal("urn:extension", asFour.Xml.Descendants().Single(element => element.Name.LocalName == "box").Name.NamespaceName);

        const string docBookFour = "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><articleinfo><author><personname><firstname>Jane</firstname><surname>Doe</surname></personname></author></articleinfo><indexterm><primary>topic</primary></indexterm></article>";
        DocBookDocument asFive = DocBookDocument.FromOfficeDocumentModel(
            DocBookDocument.Parse(docBookFour).ToOfficeDocumentModel().Value,
            profile: DocBookProfile.DocBook52).Value;
        XElement[] fiveVocabulary = asFive.Xml.Descendants()
            .Where(element => new[] { "personname", "firstname", "surname", "primary" }.Contains(element.Name.LocalName)).ToArray();
        Assert.Equal(4, fiveVocabulary.Length);
        Assert.All(fiveVocabulary,
            element => Assert.Equal(DocBookSchemaProfiles.DocBook52.NamespaceUri, element.Name.NamespaceName));
    }

    [Fact]
    public void SharedModelKeepsSectionAuthorsOutOfDocumentAttribution() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><section><title>Chapter</title><info><author>Jane Doe</author></info></section></article>";

        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;

        Assert.Null(model.Source.Author);
        Assert.Contains(model.Metadata, entry => entry.Name == "author" && entry.Value == "Jane Doe");
        OfficeDocumentModelNode section = model.Structure.Single(node => node.Kind == "section");
        Assert.Contains(section.Children.Single(node => node.Kind == "metadata").Children,
            node => node.Kind == "author" && node.Text == "Jane Doe");
    }

    [Fact]
    public void SharedConversionBoundsCumulativeSectionHeadingPaths() {
        DocBookDocument document = DocBookDocument.CreateArticle();
        DocBookNode first = document.AddSection(new string('a', 2_000));
        DocBookNode second = first.AddSection(new string('b', 2_000));
        second.AddSection(new string('c', 2_000));

        OfficeDocumentModel model = document.ToOfficeDocumentModel().Value;
        OfficeDocumentModelNode[] sections = model.Structure
            .SelectMany(root => new[] { root, root.Children.Single(node => node.Kind == "section"),
                root.Children.Single(node => node.Kind == "section").Children.Single(node => node.Kind == "section") })
            .ToArray();

        Assert.All(sections, section => Assert.True(section.Location.HeadingPath!.Length <= 1_024));
    }

    [Fact]
    public void DocBookFourUsesSectionInfoForTypedSectionMetadata() {
        DocBookDocument created = DocBookDocument.CreateArticle(DocBookProfile.DocBook45);
        DocBookNode sectionInfo = created.AddSection("Section").Add(DocBookNodeKind.Info);
        Assert.Equal("sectioninfo", sectionInfo.Name);
        Assert.Equal(DocBookNodeKind.Info, DocBookDocument.Parse(created.ToDocBook()).Root.Children
            .Single(node => node.Kind == DocBookNodeKind.Section).Children.Single(node => node.Name == "sectioninfo").Kind);

        const string docBookFive = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><section><title>Section</title><info><author>Jane Doe</author></info></section></article>";
        DocBookDocument converted = DocBookDocument.FromOfficeDocumentModel(
            DocBookDocument.Parse(docBookFive).ToOfficeDocumentModel().Value,
            profile: DocBookProfile.DocBook45).Value;
        Assert.Contains(converted.Xml.Descendants(), element => element.Name.LocalName == "sectioninfo");
    }

    [Fact]
    public void SharedModelPublishesAuthorsAndNormalizedLinks() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:xl=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><info><title>T</title><author><personname><firstname>Jane</firstname><surname>Doe</surname></personname></author></info><para><link xl:href=\"https://example.test\">Site</link><xref linkend=\"target\"/></para></article>";

        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;

        Assert.Equal("Jane Doe", model.Source.Author);
        Assert.Contains(model.Metadata, entry => entry.Name == "author" && entry.Value == "Jane Doe");
        Assert.Contains(model.Links, link => link.Uri == "https://example.test" && link.Text == "Site");
        Assert.Contains(model.Links, link => link.DestinationName == "target" && link.Kind == "cross-reference");
    }

    [Fact]
    public void NamespacedLookalikesRemainExtensionsWithoutProfileErrorsOrAliasLoss() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:extension\" version=\"5.2\"><x:ulink/><x:sect1/><x:simpara/></article>";
        DocBookDocument document = DocBookDocument.Parse(source);

        Assert.DoesNotContain(document.Validate().Diagnostics, diagnostic => diagnostic.Code == "DB014");
        Assert.DoesNotContain(document.ToOfficeDocumentModel().Diagnostics, diagnostic => diagnostic.Code == "DB115");
    }
}
