using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

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
        DocBookNode group = section.AddTable("Data").Add(DocBookNodeKind.TableGroup);
        group.SetAttribute("cols", "1");
        group.Add(DocBookNodeKind.TableBody)
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
    public void InternalSubsetProcessingInstructionsDoNotImpersonateEntityDeclarations() {
        const string source = "<!DOCTYPE article [<?audit <!ENTITY % sample SYSTEM \"uri\">?><!ENTITY safe \"ok\">]><article><para>&safe;</para></article>";

        DocBookDocument document = DocBookDocument.Parse(source);

        Assert.Equal("ok", document.Xml.Root!.Value);
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void TypedAuthorTextUsesPersonNameContent(DocBookProfile profile) {
        DocBookDocument document = DocBookDocument.CreateArticle(profile);

        DocBookNode author = document.Root.Add(DocBookNodeKind.Author, "Jane Doe");
        DocBookNode personName = Assert.Single(author.Children);

        Assert.Equal("personname", personName.Name);
        Assert.Equal("Jane Doe", personName.Text);
        XElement authorElement = document.Xml.Descendants().Single(element => element.Name.LocalName == "author");
        Assert.Equal(profile == DocBookProfile.DocBook45 ? "articleinfo" : "info", authorElement.Parent!.Name.LocalName);
        Assert.DoesNotContain(authorElement.Nodes().OfType<XText>(), text => !string.IsNullOrWhiteSpace(text.Value));
        Assert.DoesNotContain(document.Validate().Diagnostics, diagnostic => diagnostic.Code == "DB018");

        author.Text = "Janet Doe";

        Assert.Equal("Janet Doe", Assert.Single(author.Children).Text);
        Assert.Equal("personname", Assert.Single(author.Children).Name);
        Assert.DoesNotContain(document.Validate().Diagnostics, diagnostic => diagnostic.Code == "DB018");

        DocBookNode sectionAuthor = document.AddSection("Section").Add(DocBookNodeKind.Author, "Section Author");

        Assert.Equal(profile == DocBookProfile.DocBook45 ? "sectioninfo" : "info",
            document.Xml.Descendants().Single(element => element.Name.LocalName == "personname" &&
                element.Value == "Section Author").Parent!.Parent!.Name.LocalName);
        Assert.Equal("personname", Assert.Single(sectionAuthor.Children).Name);
        Assert.True(document.Validate().IsValid);
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45, "chapterinfo")]
    [InlineData(DocBookProfile.DocBook52, "info")]
    public void TypedBookComponentMetadataUsesProfileSpecificContainer(DocBookProfile profile, string expectedName) {
        DocBookDocument document = DocBookDocument.CreateBook(profile);
        document.AddParagraph("Body");
        DocBookNode chapter = document.Root.Children.Single(node => node.Name == "chapter");

        DocBookNode info = chapter.Add(DocBookNodeKind.Info);
        chapter.Add(DocBookNodeKind.Author, "Jane Doe");

        Assert.Equal(expectedName, info.Name);
        Assert.Equal(expectedName, chapter.Children.First().Name);
        Assert.Equal(expectedName, document.Xml.Descendants().Single(element =>
            element.Name.LocalName == "personname" && element.Value == "Jane Doe").Parent!.Parent!.Name.LocalName);
        Assert.True(document.Validate().IsValid);
    }

    [Fact]
    public void ValidationRejectsWrongMetadataAliasUnderBookComponent() {
        const string source = "<!DOCTYPE book PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><book><chapter><bookinfo/><title>Chapter</title></chapter></book>";

        DocBookValidationResult validation = DocBookDocument.Parse(source).Validate();

        Assert.Contains(validation.Diagnostics, diagnostic => diagnostic.Code == "DB015" &&
            diagnostic.Severity == DocBookDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void TitlelessTableUsesInformalTable(DocBookProfile profile) {
        DocBookDocument document = DocBookDocument.CreateArticle(profile);
        DocBookNode table = document.Root.AddTable();
        DocBookNode group = table.Add(DocBookNodeKind.TableGroup);
        group.SetAttribute("cols", "1");
        group.Add(DocBookNodeKind.TableBody)
            .Add(DocBookNodeKind.Row).Add(DocBookNodeKind.Entry, "Value");

        Assert.Equal("informaltable", table.Name);
        Assert.True(document.Validate().IsValid);
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45, null)]
    [InlineData(DocBookProfile.DocBook52, null)]
    [InlineData(DocBookProfile.DocBook52, "0")]
    [InlineData(DocBookProfile.DocBook52, "-1")]
    [InlineData(DocBookProfile.DocBook52, "many")]
    public void ValidationRequiresPositiveCalsColumnCount(DocBookProfile profile, string? columns) {
        DocBookDocument document = DocBookDocument.CreateArticle(profile);
        DocBookNode group = document.Root.AddTable().Add(DocBookNodeKind.TableGroup);
        if (columns != null) group.SetAttribute("cols", columns);
        group.Add(DocBookNodeKind.TableBody).Add(DocBookNodeKind.Row).Add(DocBookNodeKind.Entry, "Value");

        Assert.Contains(document.Validate().Diagnostics, diagnostic => diagnostic.Code == "DB021" &&
            diagnostic.Severity == DocBookDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData("<tbody/>", "tbody")]
    [InlineData("<thead/><tbody><row><entry>Value</entry></row></tbody>", "thead")]
    [InlineData("<tfoot/><tbody><row><entry>Value</entry></row></tbody>", "tfoot")]
    [InlineData("<tbody><row/></tbody>", "row")]
    public void ValidationRejectsEmptyCalsContainers(string tableContent, string expectedContainer) {
        string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup cols=\"1\">" +
            tableContent + "</tgroup></informaltable></article>";

        DocBookValidationResult validation = DocBookDocument.Parse(source).Validate();

        Assert.Contains(validation.Diagnostics, diagnostic => diagnostic.Code == "DB012" &&
            diagnostic.Severity == DocBookDiagnosticSeverity.Error &&
            diagnostic.Message.IndexOf(expectedContainer, StringComparison.OrdinalIgnoreCase) >= 0);
    }

    [Fact]
    public void CrossReferenceRejectsDirectTextAndValidationRejectsTextContent() {
        DocBookDocument document = DocBookDocument.CreateArticle();
        Assert.Throws<ArgumentException>(() => document.Root.Add(DocBookNodeKind.CrossReference, "Label"));
        DocBookNode crossReference = document.AddParagraph(string.Empty).Add(DocBookNodeKind.CrossReference);
        Assert.Throws<InvalidOperationException>(() => crossReference.Text = "Label");

        DocBookDocument parsed = DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><xref linkend=\"target\">Label</xref><section xml:id=\"target\"><title>Target</title></section></article>");
        Assert.Contains(parsed.Validate().Diagnostics, diagnostic => diagnostic.Code == "DB017" &&
            diagnostic.Message.IndexOf("empty", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    [Fact]
    public void ValidationRejectsDirectTextOnAuthorElements() {
        DocBookDocument document = DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><info><author>Jane Doe</author></info></article>");

        Assert.Contains(document.Validate().Diagnostics, diagnostic => diagnostic.Code == "DB018" &&
            diagnostic.Severity == DocBookDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void TypedMetadataContainersRejectDirectText(DocBookProfile profile) {
        DocBookDocument document = DocBookDocument.CreateArticle(profile);

        Assert.Throws<ArgumentException>(() => document.Root.Add(DocBookNodeKind.Info, "metadata"));
        Assert.Throws<ArgumentException>(() => document.AddSection("Section").Add(DocBookNodeKind.Info, "metadata"));
        DocBookNode info = document.Root.Add(DocBookNodeKind.Info);
        Assert.Throws<InvalidOperationException>(() => info.Text = "metadata");
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45, null)]
    [InlineData(DocBookProfile.DocBook45, "")]
    [InlineData(DocBookProfile.DocBook45, "  ")]
    [InlineData(DocBookProfile.DocBook52, null)]
    [InlineData(DocBookProfile.DocBook52, "")]
    [InlineData(DocBookProfile.DocBook52, "  ")]
    public void AddImageRejectsBlankReferencesBeforeMutation(DocBookProfile profile, string? fileReference) {
        DocBookDocument document = DocBookDocument.CreateArticle(profile);

        Assert.Throws<ArgumentException>(() => document.Root.AddImage(fileReference!));

        Assert.Empty(document.Root.Children);
    }

    [Fact]
    public void ValidationRejectsDirectTextInMetadataContainers() {
        DocBookDocument document = DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><info>metadata</info></article>");

        Assert.Contains(document.Validate().Diagnostics, diagnostic => diagnostic.Code == "DB018" &&
            diagnostic.Severity == DocBookDiagnosticSeverity.Error &&
            diagnostic.Message.IndexOf("direct text", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void ElementOnlyTypedNodesRejectDirectTextAndEmptyListsAreInvalid(DocBookProfile profile) {
        DocBookDocument document = DocBookDocument.CreateArticle(profile);
        DocBookNodeKind[] elementOnlyKinds = {
            DocBookNodeKind.Section,
            DocBookNodeKind.ItemizedList,
            DocBookNodeKind.OrderedList,
            DocBookNodeKind.VariableList,
            DocBookNodeKind.ListItem,
            DocBookNodeKind.Table,
            DocBookNodeKind.TableGroup,
            DocBookNodeKind.TableHead,
            DocBookNodeKind.TableBody,
            DocBookNodeKind.Row,
            DocBookNodeKind.Note,
            DocBookNodeKind.Tip,
            DocBookNodeKind.Important,
            DocBookNodeKind.Caution,
            DocBookNodeKind.Warning,
            DocBookNodeKind.Figure,
            DocBookNodeKind.MediaObject,
            DocBookNodeKind.ImageObject,
            DocBookNodeKind.ImageData,
            DocBookNodeKind.Caption,
            DocBookNodeKind.Index,
            DocBookNodeKind.IndexTerm
        };
        foreach (DocBookNodeKind kind in elementOnlyKinds) {
            Assert.Throws<ArgumentException>(() => document.Root.Add(kind, "orphan"));
        }
        DocBookNode list = document.Root.Add(DocBookNodeKind.ItemizedList);
        Assert.Throws<InvalidOperationException>(() => list.Text = "orphan");

        string source = profile == DocBookProfile.DocBook52
            ? "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><itemizedlist>orphan</itemizedlist><orderedlist/><variablelist/></article>"
            : "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><itemizedlist>orphan</itemizedlist><orderedlist/><variablelist/></article>";
        DocBookValidationResult validation = DocBookDocument.Parse(source).Validate();

        Assert.Contains(validation.Diagnostics, diagnostic => diagnostic.Code == "DB018" &&
            diagnostic.Message.IndexOf("itemizedlist", StringComparison.OrdinalIgnoreCase) >= 0);
        Assert.Equal(3, validation.Diagnostics.Count(diagnostic => diagnostic.Code == "DB012" &&
            diagnostic.Message.IndexOf("list", StringComparison.OrdinalIgnoreCase) >= 0));
    }

    [Fact]
    public void ValidationAllowsInfoUnderAcceptedBookComponents() {
        const string source = "<book xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><chapter><info><title>Chapter</title></info><para>Body</para></chapter></book>";

        DocBookValidationResult validation = DocBookDocument.Parse(source).Validate();

        Assert.True(validation.IsValid);
        Assert.DoesNotContain(validation.Diagnostics, diagnostic => diagnostic.Code == "DB015");
    }

    [Fact]
    public void ValidationRejectsListItemsOutsideSupportedListParents() {
        DocBookDocument invalid = DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><listitem><para>Root item</para></listitem><section><title>Section</title><listitem><para>Section item</para></listitem></section></article>");

        Assert.Equal(2, invalid.Validate().Diagnostics.Count(diagnostic => diagnostic.Code == "DB015"));

        DocBookDocument variableList = DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><variablelist><varlistentry><term>Name</term><listitem><para>Value</para></listitem></varlistentry></variablelist></article>");
        Assert.DoesNotContain(variableList.Validate().Diagnostics, diagnostic => diagnostic.Code == "DB015");
    }

    [Theory]
    [InlineData("<book xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para>Body</para></book>")]
    [InlineData("<!DOCTYPE book PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><book><section><title>Body</title></section></book>")]
    public void BoundedValidationRejectsTypedBlocksDirectlyUnderBook(string source) {
        DocBookValidationResult validation = DocBookDocument.Parse(source).Validate();

        Assert.False(validation.IsValid);
        Assert.Contains(validation.Diagnostics, diagnostic =>
            diagnostic.Code == "DB015" && diagnostic.Severity == DocBookDiagnosticSeverity.Error);
    }

    [Fact]
    public void BoundedBookValidationDoesNotTreatForeignExtensionsAsTypedRootContent() {
        DocBookValidationResult validation = DocBookDocument.Parse(
            "<book xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:test\" version=\"5.2\"><x:component/></book>").Validate();

        Assert.DoesNotContain(validation.Diagnostics, diagnostic => diagnostic.Code == "DB015");
        Assert.Contains(validation.Diagnostics, diagnostic => diagnostic.Code == "DB010");
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void TypedBookAdditionsUseATitledChapterContainer(DocBookProfile profile) {
        DocBookDocument document = DocBookDocument.CreateBook(profile);

        document.AddParagraph("Body");

        XElement chapter = Assert.Single(document.Xml.Root!.Elements(), element => element.Name.LocalName == "chapter");
        Assert.Equal("Content", chapter.Elements().First().Value);
        Assert.Equal("title", chapter.Elements().First().Name.LocalName);
        Assert.Equal("Body", chapter.Elements().Last().Value);
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
        OfficeDocumentModelNode chapter = Assert.Single(model.Structure, node => node.Kind == "extension:chapter");
        Assert.Contains(chapter.Children, node => node.Kind == "section");
        Assert.DoesNotContain(document.Xml.Root!.Elements(), element =>
            element.Name.LocalName == "section" || element.Name.LocalName == "para");
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

    [Fact]
    public void SharedConversionReportsNamespacedRootVersionAsAnExtensionAttribute() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:extension\" version=\"5.2\" x:version=\"producer\"><para>P</para></article>";

        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(source).ToOfficeDocumentModel();

        Assert.True(converted.HasLoss);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB106");
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
    public void SharedTableProjectionMovesDtdOrderedFooterRowsAfterBodyRows() {
        const string source = "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><informaltable><tgroup cols=\"1\"><tfoot><row><entry>Total</entry></row></tfoot><tbody><row><entry>Body</entry></row></tbody></tgroup></informaltable></article>";

        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(source).ToOfficeDocumentModel();
        OfficeDocumentModelTable table = Assert.Single(converted.Value.Tables);

        Assert.Equal(new[] { "Body", "Total" }, table.Rows.Select(row => row.Single()).ToArray());
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB119");
    }

    [Fact]
    public void SharedTableProjectionPrioritizesBodyRowsBeforeFooterRowsAtTheRowLimit() {
        const string source = "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><informaltable><tgroup cols=\"1\"><tfoot><row><entry>Total</entry></row></tfoot><tbody><row><entry>Body</entry></row></tbody></tgroup></informaltable></article>";

        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(source).ToOfficeDocumentModel(
            options: new DocBookConversionOptions { MaxTableRows = 1 });
        OfficeDocumentModelTable table = Assert.Single(converted.Value.Tables);

        Assert.Equal("Body", Assert.Single(Assert.Single(table.Rows)));
        Assert.Equal(2, table.TotalRowCount);
        Assert.True(table.Truncated);
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
        Assert.Equal("Outer", Assert.Single(Assert.Single(table.Rows)));
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
    public void SharedProjectionUsesTableAndFigureTitlesInRecursiveHeadingPaths() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><section><title>Section</title><table><title>Values</title><tgroup cols=\"1\"><tbody><row><entry>A</entry></row></tbody></tgroup></table><figure><title>Diagram</title><mediaobject><imageobject><imagedata fileref=\"diagram.png\"/></imageobject></mediaobject></figure></section></article>";

        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;

        OfficeDocumentModelNode table = FindStructureNode(model.Structure, "table");
        OfficeDocumentModelNode figure = FindStructureNode(model.Structure, "figure");
        Assert.Equal("Section / Values", table.Location.HeadingPath);
        Assert.Equal(Assert.Single(model.Tables).Location!.HeadingPath, table.Location.HeadingPath);
        Assert.Equal("Section / Diagram", figure.Location.HeadingPath);
        Assert.Equal(figure.Location.HeadingPath, Assert.Single(model.Assets).Location.HeadingPath);
    }

    [Fact]
    public void SharedModelHeadingPathsIncludeTitledBookComponents() {
        const string source = "<book xmlns=\"http://docbook.org/ns/docbook\" xmlns:xl=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><chapter><title>Alpha</title><section><title>Details</title><para><link xl:href=\"https://example.test/alpha\">Alpha link</link></para></section></chapter><chapter><title>Beta</title><section><title>Details</title><mediaobject><imageobject><imagedata fileref=\"beta.png\"/></imageobject></mediaobject><table><title>Values</title><tgroup><tbody><row><entry>B</entry></row></tbody></tgroup></table></section></chapter></book>";

        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        OfficeDocumentModelNode[] chapters = model.Structure
            .Where(node => node.Kind.EndsWith("chapter", StringComparison.Ordinal)).ToArray();
        OfficeDocumentModelNode[] sections = chapters
            .Select(chapter => chapter.Children.Single(node => node.Kind == "section")).ToArray();

        Assert.Equal(new[] { "Alpha", "Beta" }, chapters.Select(chapter => chapter.Location.HeadingPath));
        Assert.Equal(new[] { "Alpha / Details", "Beta / Details" }, sections.Select(section => section.Location.HeadingPath));
        Assert.Equal("Alpha / Details", Assert.Single(model.Links).Location!.HeadingPath);
        Assert.Equal("Beta / Details", Assert.Single(model.Assets).Location!.HeadingPath);
        Assert.Equal("Beta / Details / Values", Assert.Single(model.Tables).Location!.HeadingPath);
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

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void BookToArticleConversionNormalizesBookOnlyComponentsToSections(DocBookProfile targetProfile) {
        const string source = "<book xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><part><title>Part</title><chapter><title>Chapter</title><para>Body</para></chapter></part><appendix><title>Appendix</title><para>Tail</para></appendix></book>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(
            model, DocBookDocumentKind.Article, targetProfile);

        string[] bookOnlyNames = { "chapter", "part", "appendix", "preface", "reference", "article" };
        Assert.DoesNotContain(converted.Value.Xml.Root!.Descendants(), element =>
            bookOnlyNames.Contains(element.Name.LocalName, StringComparer.Ordinal));
        Assert.Equal(new[] { "Part", "Chapter", "Appendix" }, converted.Value.Xml.Descendants()
            .Where(element => element.Name.LocalName == "section")
            .Select(element => element.Elements().First(child => child.Name.LocalName == "title").Value));
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB126");
        Assert.True(converted.Value.Validate().IsValid);

        string invalidArticle = targetProfile == DocBookProfile.DocBook52
            ? "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><chapter><title>Invalid</title></chapter></article>"
            : "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><chapter><title>Invalid</title></chapter></article>";
        Assert.Contains(DocBookDocument.Parse(invalidArticle).Validate().Diagnostics, diagnostic =>
            diagnostic.Code == "DB015" && diagnostic.Severity == DocBookDiagnosticSeverity.Error);
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
    public void ProfileConversionDropsStaleDefaultNamespaceDeclarationsOnTypedVocabulary() {
        const string docBookFive = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><section xmlns=\"http://docbook.org/ns/docbook\"><title>Five</title></section></article>";
        DocBookDocument asFour = DocBookDocument.FromOfficeDocumentModel(
            DocBookDocument.Parse(docBookFive).ToOfficeDocumentModel().Value,
            profile: DocBookProfile.DocBook45).Value;
        XElement fourSection = asFour.Xml.Descendants().Single(element => element.Name.LocalName == "section");

        Assert.Equal(XNamespace.None, fourSection.Name.Namespace);
        Assert.DoesNotContain(fourSection.Attributes(), attribute => attribute.IsNamespaceDeclaration &&
            attribute.Value == DocBookSchemaProfiles.DocBook52.NamespaceUri);

        const string docBookFour = "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><section xmlns=\"\"><title>Four</title></section></article>";
        DocBookDocument asFive = DocBookDocument.FromOfficeDocumentModel(
            DocBookDocument.Parse(docBookFour).ToOfficeDocumentModel().Value,
            profile: DocBookProfile.DocBook52).Value;
        XElement fiveSection = asFive.Xml.Descendants().Single(element => element.Name.LocalName == "section");

        Assert.Equal(DocBookSchemaProfiles.DocBook52.NamespaceUri, fiveSection.Name.NamespaceName);
        Assert.DoesNotContain(fiveSection.Attributes(), attribute =>
            attribute.IsNamespaceDeclaration && attribute.Value.Length == 0);
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void AddLinkWrapsBlockContainersAndRejectsMetadataContainers(DocBookProfile profile) {
        DocBookDocument article = DocBookDocument.CreateArticle(profile);
        article.Root.AddLink("Root", "https://example.test/root");
        article.AddSection("Section").AddLink("Section", "https://example.test/section");

        Assert.All(article.Xml.Descendants().Where(element => element.Name.LocalName == "link" || element.Name.LocalName == "ulink"),
            link => Assert.Equal("para", link.Parent!.Name.LocalName));
        Assert.True(article.Validate().IsValid);
        Assert.Throws<InvalidOperationException>(() =>
            article.Root.Add(DocBookNodeKind.Info).AddLink("Metadata", "https://example.test/metadata"));

        DocBookDocument book = DocBookDocument.CreateBook(profile);
        book.Root.AddLink("Book", "https://example.test/book");

        XElement bookLink = book.Xml.Descendants().Single(element => element.Name.LocalName == "link" || element.Name.LocalName == "ulink");
        Assert.Equal("para", bookLink.Parent!.Name.LocalName);
        Assert.Equal("chapter", bookLink.Parent.Parent!.Name.LocalName);
        Assert.True(book.Validate().IsValid);
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void TypedLinksRequireTargetsAndLinkHelpersRejectBlankTargets(DocBookProfile profile) {
        DocBookDocument document = DocBookDocument.CreateArticle(profile);
        DocBookNode paragraph = document.AddParagraph("Body");
        paragraph.Add(DocBookNodeKind.Link, "Site");

        Assert.Contains(document.Validate().Diagnostics, diagnostic =>
            diagnostic.Code == "DB017" && diagnostic.Severity == DocBookDiagnosticSeverity.Error);
        Assert.Throws<ArgumentException>(() => paragraph.AddLink("Site", "  "));
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void InlineContainersRejectTypedBlockChildrenAndValidationCatchesParsedPlacement(DocBookProfile profile) {
        DocBookDocument document = DocBookDocument.CreateArticle(profile);
        Assert.Throws<InvalidOperationException>(() => document.AddParagraph("Body").AddSection("Nested"));
        Assert.Throws<InvalidOperationException>(() => document.AddSection("Section").Children
            .Single(child => child.Kind == DocBookNodeKind.Title).AddParagraph("Nested"));

        string source = profile == DocBookProfile.DocBook52
            ? "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para><section><title>Nested</title></section></para></article>"
            : "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><para><section><title>Nested</title></section></para></article>";
        Assert.Contains(DocBookDocument.Parse(source).Validate().Diagnostics, diagnostic =>
            diagnostic.Code == "DB015" && diagnostic.Severity == DocBookDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void ElementOnlyTypedContainersRejectUnsupportedChildrenBeforeMutation(DocBookProfile profile) {
        DocBookDocument document = DocBookDocument.CreateArticle(profile);
        DocBookNode table = document.Root.AddTable("Values");
        DocBookNode group = table.Add(DocBookNodeKind.TableGroup);
        group.SetAttribute("cols", "1");
        DocBookNode row = group.Add(DocBookNodeKind.TableBody).Add(DocBookNodeKind.Row);
        row.Add(DocBookNodeKind.Entry, "A");
        DocBookNode list = document.Root.AddItemizedList();
        list.AddListItem("Item");

        Assert.Throws<InvalidOperationException>(() => row.AddParagraph("Orphan"));
        Assert.Throws<InvalidOperationException>(() => table.AddParagraph("Orphan"));
        Assert.Throws<InvalidOperationException>(() => list.AddParagraph("Orphan"));

        Assert.DoesNotContain(document.Xml.Descendants(), element => element.Value == "Orphan");
        Assert.True(document.Validate().IsValid);

        string source = profile == DocBookProfile.DocBook52
            ? "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><para>Orphan</para><tgroup cols=\"1\"><tbody><row><entry>A</entry><para>Orphan</para></row></tbody></tgroup></informaltable><itemizedlist><listitem><para>Item</para></listitem><para>Orphan</para></itemizedlist></article>"
            : "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><informaltable><para>Orphan</para><tgroup cols=\"1\"><tbody><row><entry>A</entry><para>Orphan</para></row></tbody></tgroup></informaltable><itemizedlist><listitem><para>Item</para></listitem><para>Orphan</para></itemizedlist></article>";
        Assert.Equal(3, DocBookDocument.Parse(source).Validate().Diagnostics.Count(diagnostic =>
            diagnostic.Code == "DB015" && diagnostic.Severity == DocBookDiagnosticSeverity.Error));
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45, "table")]
    [InlineData(DocBookProfile.DocBook45, "informaltable")]
    [InlineData(DocBookProfile.DocBook52, "table")]
    [InlineData(DocBookProfile.DocBook52, "informaltable")]
    public void ValidationRejectsTablesWithoutTableGroups(DocBookProfile profile, string tableName) {
        string title = tableName == "table" ? "<title>Values</title>" : string.Empty;
        string source = profile == DocBookProfile.DocBook52
            ? $"<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><{tableName}>{title}</{tableName}></article>"
            : $"<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><{tableName}>{title}</{tableName}></article>";

        DocBookValidationResult validation = DocBookDocument.Parse(source).Validate();

        Assert.Contains(validation.Diagnostics, diagnostic => diagnostic.Code == "DB012" &&
            diagnostic.Severity == DocBookDiagnosticSeverity.Error &&
            diagnostic.Message.IndexOf("tgroup", StringComparison.Ordinal) >= 0);
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void TypedTableHeadersAreInsertedBeforeBodies(DocBookProfile profile) {
        DocBookDocument document = DocBookDocument.CreateArticle(profile);
        DocBookNode group = document.Root.AddTable().Add(DocBookNodeKind.TableGroup);
        group.SetAttribute("cols", "1");
        group.Add(DocBookNodeKind.TableBody).Add(DocBookNodeKind.Row).Add(DocBookNodeKind.Entry, "Body");
        group.Add(DocBookNodeKind.TableHead).Add(DocBookNodeKind.Row).Add(DocBookNodeKind.Entry, "Heading");

        Assert.Equal(new[] { "thead", "tbody" }, group.Children.Select(child => child.Name));
        Assert.True(document.Validate().IsValid);
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45, "thead")]
    [InlineData(DocBookProfile.DocBook45, "tfoot")]
    [InlineData(DocBookProfile.DocBook45, "colspec")]
    [InlineData(DocBookProfile.DocBook45, "spanspec")]
    [InlineData(DocBookProfile.DocBook52, "thead")]
    [InlineData(DocBookProfile.DocBook52, "tfoot")]
    [InlineData(DocBookProfile.DocBook52, "colspec")]
    [InlineData(DocBookProfile.DocBook52, "spanspec")]
    public void ValidationRejectsOutOfOrderCalsSections(DocBookProfile profile, string lateSection) {
        string lateContent = lateSection == "thead" || lateSection == "tfoot"
            ? "<row><entry>Late</entry></row>"
            : string.Empty;
        string source = profile == DocBookProfile.DocBook52
            ? $"<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup cols=\"1\"><tbody><row><entry>Body</entry></row></tbody><{lateSection}>{lateContent}</{lateSection}></tgroup></informaltable></article>"
            : $"<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><informaltable><tgroup cols=\"1\"><tbody><row><entry>Body</entry></row></tbody><{lateSection}>{lateContent}</{lateSection}></tgroup></informaltable></article>";

        DocBookValidationResult validation = DocBookDocument.Parse(source).Validate();

        Assert.Contains(validation.Diagnostics, diagnostic => diagnostic.Code == "DB020" &&
            diagnostic.Severity == DocBookDiagnosticSeverity.Error &&
            diagnostic.Message.IndexOf(lateSection, StringComparison.Ordinal) >= 0);
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void ValidationRejectsDuplicateCalsFooters(DocBookProfile profile) {
        const string footers = "<tfoot><row><entry>One</entry></row></tfoot><tfoot><row><entry>Two</entry></row></tfoot>";
        string source = profile == DocBookProfile.DocBook52
            ? $"<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup cols=\"1\">{footers}<tbody><row><entry>Body</entry></row></tbody></tgroup></informaltable></article>"
            : $"<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><informaltable><tgroup cols=\"1\">{footers}<tbody><row><entry>Body</entry></row></tbody></tgroup></informaltable></article>";

        Assert.Contains(DocBookDocument.Parse(source).Validate().Diagnostics, diagnostic =>
            diagnostic.Code == "DB019" && diagnostic.Severity == DocBookDiagnosticSeverity.Error &&
            diagnostic.Message.IndexOf("tfoot", StringComparison.Ordinal) >= 0);
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void TypedSingletonChildrenRejectDuplicatesAndValidationCatchesParsedDuplicates(DocBookProfile profile) {
        DocBookDocument document = DocBookDocument.CreateArticle(profile);
        DocBookNode section = document.AddSection("Section");
        section.Add(DocBookNodeKind.Subtitle, "Subtitle");
        DocBookNode group = section.AddTable().Add(DocBookNodeKind.TableGroup);
        group.SetAttribute("cols", "1");
        group.Add(DocBookNodeKind.TableHead).Add(DocBookNodeKind.Row).Add(DocBookNodeKind.Entry, "Heading");
        group.Add(DocBookNodeKind.TableBody).Add(DocBookNodeKind.Row).Add(DocBookNodeKind.Entry, "A");
        DocBookNode media = section.AddImage("image.png", "Caption");
        DocBookNode imageObject = media.Children.Single(child => child.Kind == DocBookNodeKind.ImageObject);

        Assert.Throws<InvalidOperationException>(() => section.Add(DocBookNodeKind.Title, "Duplicate"));
        Assert.Throws<InvalidOperationException>(() => section.Add(DocBookNodeKind.Subtitle, "Duplicate"));
        Assert.Throws<InvalidOperationException>(() => group.Add(DocBookNodeKind.TableHead));
        Assert.Throws<InvalidOperationException>(() => group.Add(DocBookNodeKind.TableBody));
        Assert.Throws<InvalidOperationException>(() => imageObject.Add(DocBookNodeKind.ImageData));
        Assert.Throws<InvalidOperationException>(() => media.Add(DocBookNodeKind.Caption));
        Assert.True(document.Validate().IsValid);

        string source = profile == DocBookProfile.DocBook52
            ? "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><section><title>Section</title><subtitle>One</subtitle><subtitle>Two</subtitle></section><informaltable><tgroup cols=\"1\"><tbody><row><entry>A</entry></row></tbody><tbody><row><entry>B</entry></row></tbody></tgroup></informaltable></article>"
            : "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><section><title>Section</title><subtitle>One</subtitle><subtitle>Two</subtitle></section><informaltable><tgroup cols=\"1\"><tbody><row><entry>A</entry></row></tbody><tbody><row><entry>B</entry></row></tbody></tgroup></informaltable></article>";
        DocBookValidationResult validation = DocBookDocument.Parse(source).Validate();

        Assert.Equal(2, validation.Diagnostics.Count(diagnostic => diagnostic.Code == "DB019" &&
            diagnostic.Severity == DocBookDiagnosticSeverity.Error));
    }

    [Fact]
    public void ValidationRejectsInlineLinksOutsideInlineContentParents() {
        DocBookDocument document = DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:xl=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><link xl:href=\"https://example.test/\">Site</link></article>");

        Assert.Contains(document.Validate().Diagnostics, diagnostic => diagnostic.Code == "DB015" &&
            diagnostic.Severity == DocBookDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData("emphasis")]
    [InlineData("phrase")]
    public void ValidationAllowsLinksInUntypedInlineVocabulary(string inlineName) {
        DocBookDocument document = DocBookDocument.Parse(
            $"<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:xl=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><para><{inlineName}><link xl:href=\"https://example.test/\">Site</link></{inlineName}></para></article>");

        DocBookValidationResult validation = document.Validate();

        Assert.True(validation.IsValid);
        Assert.DoesNotContain(validation.Diagnostics, diagnostic => diagnostic.Code == "DB015");
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void TypedMetadataCreationReusesOneContainerAndValidationRejectsDuplicates(DocBookProfile profile) {
        DocBookDocument article = DocBookDocument.CreateArticle(profile);
        article.Root.Add(DocBookNodeKind.Info);
        article.Root.Add(DocBookNodeKind.Info);
        DocBookNode section = article.AddSection("Section");
        section.Add(DocBookNodeKind.Info);
        section.Add(DocBookNodeKind.Info);

        DocBookDocument book = DocBookDocument.CreateBook(profile);
        book.Root.Add(DocBookNodeKind.Info);
        book.Root.Add(DocBookNodeKind.Info);
        book.AddParagraph("Body");
        DocBookNode chapter = book.Root.Children.Single(node => node.Name == "chapter");
        chapter.Add(DocBookNodeKind.Info);
        chapter.Add(DocBookNodeKind.Info);

        Assert.Single(article.Xml.Root!.Elements(), element => element.Name.LocalName.EndsWith("info", StringComparison.Ordinal));
        Assert.Single(section.Children, child => child.Name.EndsWith("info", StringComparison.Ordinal));
        Assert.Single(book.Xml.Root!.Elements(), element => element.Name.LocalName.EndsWith("info", StringComparison.Ordinal));
        Assert.Single(chapter.Children, child => child.Name.EndsWith("info", StringComparison.Ordinal));

        XElement chapterElement = book.Xml.Descendants().Single(element => element.Name.LocalName == "chapter");
        XElement chapterInfo = chapterElement.Elements().Single(element => element.Name.LocalName.EndsWith("info", StringComparison.Ordinal));
        chapterElement.Add(new XElement(chapterInfo.Name));

        Assert.Contains(book.Validate().Diagnostics, diagnostic => diagnostic.Code == "DB019" &&
            diagnostic.Severity == DocBookDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void TypedTitlesAndSubtitlesAreInsertedBeforeBodyAndValidationRejectsLateParsedHeaders(DocBookProfile profile) {
        DocBookDocument article = DocBookDocument.CreateArticle(profile);
        DocBookNode section = article.Root.Add(DocBookNodeKind.Section);
        section.AddParagraph("Body");
        section.Add(DocBookNodeKind.Subtitle, "Late subtitle");
        section.Add(DocBookNodeKind.Title, "Late section");
        Assert.Equal(new[] { "title", "subtitle", "para" }, section.Children.Select(child => child.Name));

        DocBookNode table = article.Root.Add(DocBookNodeKind.Table);
        DocBookNode tableGroup = table.Add(DocBookNodeKind.TableGroup);
        tableGroup.SetAttribute("cols", "1");
        tableGroup.Add(DocBookNodeKind.TableBody)
            .Add(DocBookNodeKind.Row).Add(DocBookNodeKind.Entry, "Value");
        table.Add(DocBookNodeKind.Title, "Late table");
        Assert.Equal("title", table.Children.First().Name);

        DocBookDocument book = DocBookDocument.CreateBook(profile);
        book.AddParagraph("Book body");
        DocBookNode chapter = book.Root.Children.Single(child => child.Name == "chapter");
        chapter.Children.Single(child => child.Kind == DocBookNodeKind.Title).Remove();
        chapter.Add(DocBookNodeKind.Subtitle, "Late subtitle");
        chapter.Add(DocBookNodeKind.Title, "Late chapter");
        Assert.Equal(new[] { "title", "subtitle", "para" }, chapter.Children.Select(child => child.Name));

        Assert.True(article.Validate().IsValid);
        Assert.True(book.Validate().IsValid);

        string parsedSource = profile == DocBookProfile.DocBook52
            ? "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><section><para>Body</para><title>Late</title></section></article>"
            : "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><section><para>Body</para><title>Late</title></section></article>";
        Assert.Contains(DocBookDocument.Parse(parsedSource).Validate().Diagnostics, diagnostic =>
            diagnostic.Code == "DB020" && diagnostic.Severity == DocBookDiagnosticSeverity.Error);

        string parsedSubtitleSource = profile == DocBookProfile.DocBook52
            ? "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><section><title>Title</title><para>Body</para><subtitle>Late</subtitle></section></article>"
            : "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><section><title>Title</title><para>Body</para><subtitle>Late</subtitle></section></article>";
        Assert.Contains(DocBookDocument.Parse(parsedSubtitleSource).Validate().Diagnostics, diagnostic =>
            diagnostic.Code == "DB022" && diagnostic.Severity == DocBookDiagnosticSeverity.Error);
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
    public void LoadedBomlessUtf32SourceIsReturnedAsTextWithoutUtf8Fallback() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para>Caf\u00e9</para></article>";
        byte[] bytes = new UTF32Encoding(true, false, true).GetBytes(source);

        DocBookDocument document = DocBookDocument.Load(new MemoryStream(bytes));

        Assert.Equal(source, document.ToDocBook());
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
        Assert.Single(table.Rows);
        Assert.All(table.Rows, row => Assert.Equal(4, row.Count));
        Assert.Equal(3, table.TotalRowCount);
        Assert.True(table.Truncated);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB113");
        Assert.Throws<ArgumentOutOfRangeException>(() => DocBookDocument.Parse(source)
            .ToOfficeDocumentModel(options: new DocBookConversionOptions { MaxTableCells = 0 }));
    }

    [Fact]
    public void SharedTableProjectionBoundsCellSlotsAcrossAllTables() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup cols=\"2\"><tbody><row><entry>A</entry><entry>1</entry></row></tbody></tgroup></informaltable><informaltable><tgroup cols=\"2\"><tbody><row><entry>B</entry><entry>2</entry></row></tbody></tgroup></informaltable></article>";

        DocBookConversionResult<OfficeDocumentModel> converted = DocBookDocument.Parse(source)
            .ToOfficeDocumentModel(options: new DocBookConversionOptions { MaxTableColumns = 2, MaxTableRows = 10, MaxTableCells = 4 });

        Assert.Equal(2, converted.Value.Tables.Count);
        Assert.True(converted.Value.Tables.Sum(table => table.Columns.Count + table.Rows.Sum(row => row.Count)) <= 4);
        Assert.Empty(converted.Value.Tables[1].Columns);
        Assert.Empty(converted.Value.Tables[1].Rows);
        Assert.True(converted.Value.Tables[1].Truncated);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB113");
    }

    [Fact]
    public void SharedProjectionExcludesIndexTermsFromDisplayedText() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para>Body<indexterm><primary>topic</primary></indexterm></para></article>";

        OfficeDocumentModelNode paragraph = Assert.Single(DocBookDocument.Parse(source).ToOfficeDocumentModel().Value.Structure);

        Assert.Equal("Body", paragraph.Text);
        Assert.Contains(paragraph.Children, child => child.Kind == "index-term");

        const string extensionSource = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:extension\" version=\"5.2\"><para>Body<x:indexterm>visible</x:indexterm></para></article>";
        OfficeDocumentModelNode extensionParagraph = Assert.Single(
            DocBookDocument.Parse(extensionSource).ToOfficeDocumentModel().Value.Structure);
        Assert.Equal("Bodyvisible", extensionParagraph.Text);
    }

    [Theory]
    [InlineData("<xref/>")]
    [InlineData("<xref linkend=\"  \"/>")]
    public void BoundedValidationRejectsCrossReferencesWithoutTargets(string crossReference) {
        DocBookDocument document = DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para>" + crossReference + "</para></article>");

        Assert.Contains(document.Validate().Diagnostics, diagnostic =>
            diagnostic.Code == "DB017" && diagnostic.Severity == DocBookDiagnosticSeverity.Error);
    }

    [Fact]
    public void BoundedValidationRejectsUnsupportedCrossReferenceTargetAttributes() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:xl=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><para><xref linkend=\"target\" xl:href=\"https://example.test\"/></para></article>";

        Assert.Contains(DocBookDocument.Parse(source).Validate().Diagnostics, diagnostic =>
            diagnostic.Code == "DB016" && diagnostic.Severity == DocBookDiagnosticSeverity.Error);
    }

    [Fact]
    public void ValidationAndProjectionObserveCancellation() {
        DocBookDocument document = DocBookDocument.Parse(
            "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para>Body</para></article>");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => document.Validate(cancellationToken: cancellation.Token));
        Assert.Throws<OperationCanceledException>(() => document.ToOfficeDocumentModel(cancellationToken: cancellation.Token));
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

    [Fact]
    public void SharedProjectionIgnoresExtensionElementsThatImitateAuthorNameFields() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:extension\" version=\"5.2\"><info><author><x:personname>Internal</x:personname><personname><x:firstname>Hidden</x:firstname><firstname>Jane</firstname><surname>Doe</surname></personname><x:surname>Secret</x:surname></author></info></article>";

        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;

        Assert.Equal("Jane Doe", model.Source.Author);
        Assert.Contains(model.Metadata, entry => entry.Name == "author" && entry.Value == "Jane Doe");
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
        XElement crossReference = Assert.Single(converted.Value.Xml.Descendants(), element => element.Name.LocalName == "xref");

        Assert.Equal("Jane Doe", info.Descendants().Single(element => element.Name.LocalName == "author").Value);
        Assert.Equal(new[] { "Name", "Value", "A", "1" },
            table.Descendants().Where(element => element.Name.LocalName == "entry").Select(element => element.Value));
        Assert.Equal("assets/figure.png", image.Attribute("fileref")!.Value);
        Assert.Contains(links, link => (string?)link.Attribute(XName.Get("href", "http://www.w3.org/1999/xlink")) == "https://example.test/");
        Assert.Equal("target-id", (string?)crossReference.Attribute("linkend"));
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB103");
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB120" &&
            diagnostic.Message.IndexOf("target", StringComparison.OrdinalIgnoreCase) >= 0);
        OfficeDocumentModel projected = converted.Value.ToOfficeDocumentModel().Value;
        Assert.Contains(projected.Links, link => link.Uri == "https://example.test/" && link.Text == "Site");
        Assert.Contains(projected.Links, link => link.DestinationName == "target-id" &&
            link.Kind == "cross-reference" && string.IsNullOrEmpty(link.Text));
    }

    [Fact]
    public void SharedConversionPrefersPortableFileNamesForGenericImageAssets() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Pdf,
            Assets = new[] {
                new OfficeDocumentModelAsset {
                    Id = "figure",
                    Kind = "image",
                    FileName = "figure.png",
                    SourceObjectId = "pdf-object-12"
                }
            }
        };

        DocBookDocument converted = DocBookDocument.FromOfficeDocumentModel(model).Value;

        XElement image = Assert.Single(converted.Xml.Descendants(), element => element.Name.LocalName == "imagedata");
        Assert.Equal("figure.png", image.Attribute("fileref")!.Value);
    }

    [Theory]
    [InlineData("4.5")]
    [InlineData("5.2")]
    public void SharedConversionPreservesDistinctImageCaptionAndAlternativeText(string profile) {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Pdf,
            Metadata = new[] { new OfficeDocumentModelMetadataEntry { Category = "docbook", Name = "profile", Value = profile } },
            Assets = new[] {
                new OfficeDocumentModelAsset {
                    Id = "figure",
                    Kind = "image",
                    FileName = "figure.png",
                    Title = "Visible caption",
                    AltText = "Alternative description"
                }
            }
        };

        DocBookDocument converted = DocBookDocument.FromOfficeDocumentModel(model).Value;
        XElement media = Assert.Single(converted.Xml.Descendants(), element => element.Name.LocalName == "mediaobject");
        OfficeDocumentModelAsset restored = Assert.Single(converted.ToOfficeDocumentModel().Value.Assets);

        Assert.Equal("Alternative description", media.Elements().Single(element => element.Name.LocalName == "textobject").Value);
        Assert.Equal("Visible caption", media.Elements().Single(element => element.Name.LocalName == "caption").Value);
        Assert.Equal("Alternative description", restored.AltText);
        Assert.Equal("Visible caption", restored.Title);
    }

    [Fact]
    public void SharedConversionDiagnosesUnsupportedPagesFormsAndVisuals() {
        var model = new OfficeDocumentModel {
            Pages = new[] { new OfficeDocumentModelPage { Number = 3, Name = "Page three" } },
            Forms = new[] { new OfficeDocumentModelFormField { Id = "accept", Name = "Accept", Kind = "checkbox" } },
            Visuals = new[] { new OfficeDocumentModelVisual { Kind = "diagram", SourceName = "Flow", Content = "A -> B" } }
        };

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.True(converted.HasLoss);
        Assert.Equal(3, converted.Diagnostics.Count(diagnostic => diagnostic.Code == "DB124"));
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB124" && diagnostic.Message.IndexOf("Page three", StringComparison.Ordinal) >= 0);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB124" && diagnostic.Message.IndexOf("Accept", StringComparison.Ordinal) >= 0);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB124" && diagnostic.Message.IndexOf("Flow", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void SharedConversionDiagnosesPortableOnlyBodyChannels() {
        var model = new OfficeDocumentModel { Markdown = "# Heading", Html = "<h1>Heading</h1>" };

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.True(converted.HasLoss);
        Assert.Empty(converted.Value.Xml.Root!.Descendants());
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB124" &&
            diagnostic.Message.IndexOf("Markdown", StringComparison.Ordinal) >= 0);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB124" &&
            diagnostic.Message.IndexOf("HTML", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void SharedConversionDiagnosesUnsupportedSourceSubjectAndKeywords() {
        var model = new OfficeDocumentModel {
            Source = new OfficeDocumentModelSource { Subject = "Quarterly report", Keywords = "alpha, beta" }
        };

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.True(converted.HasLoss);
        Assert.Empty(converted.Value.Xml.Root!.Descendants());
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB125" &&
            diagnostic.Message.IndexOf("Source.Subject", StringComparison.Ordinal) >= 0);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB125" &&
            diagnostic.Message.IndexOf("Source.Keywords", StringComparison.Ordinal) >= 0);
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

    [Theory]
    [InlineData(DocBookProfile.DocBook45, false)]
    [InlineData(DocBookProfile.DocBook45, true)]
    [InlineData(DocBookProfile.DocBook52, false)]
    [InlineData(DocBookProfile.DocBook52, true)]
    public void SharedReverseConversionDoesNotResurrectDeletedRecursiveProjections(
        DocBookProfile profile,
        bool retainUnrelatedStructure) {
        string source = profile == DocBookProfile.DocBook52
            ? "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><para>Deleted block</para><informaltable><tgroup cols=\"1\"><tbody><row><entry>Deleted cell</entry></row></tbody></tgroup></informaltable><mediaobject><imageobject><imagedata fileref=\"deleted.png\"/></imageobject></mediaobject><para><link xlink:href=\"https://example.test/deleted\">Deleted link</link></para></article>"
            : "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><para>Deleted block</para><informaltable><tgroup cols=\"1\"><tbody><row><entry>Deleted cell</entry></row></tbody></tgroup></informaltable><mediaobject><imageobject><imagedata fileref=\"deleted.png\"/></imageobject></mediaobject><para><ulink url=\"https://example.test/deleted\">Deleted link</ulink></para></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        model.Structure = retainUnrelatedStructure
            ? new[] { new OfficeDocumentModelNode { Id = "retained", Kind = "paragraph", Text = "Retained" } }
            : Array.Empty<OfficeDocumentModelNode>();

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.DoesNotContain("Deleted", converted.Value.Xml.Root!.Value, StringComparison.Ordinal);
        Assert.DoesNotContain(converted.Value.Xml.Descendants(), element =>
            element.Name.LocalName == "table" || element.Name.LocalName == "informaltable" ||
            element.Name.LocalName == "imagedata" || element.Name.LocalName == "link" || element.Name.LocalName == "ulink");
        Assert.Equal(retainUnrelatedStructure, converted.Value.Xml.Descendants().Any(element =>
            element.Name.LocalName == "para" && element.Value == "Retained"));
        Assert.DoesNotContain(converted.Diagnostics, diagnostic => diagnostic.Code == "DB122" || diagnostic.Code == "DB103");
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void SharedRoundTripMatchesTitledTablesWithAlreadyQualifiedNodePaths(DocBookProfile profile) {
        string source = profile == DocBookProfile.DocBook52
            ? "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><section><title>Section</title><table><title>Values</title><tgroup cols=\"1\"><tbody><row><entry>A</entry></row></tbody></tgroup></table></section></article>"
            : "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><section><title>Section</title><table><title>Values</title><tgroup cols=\"1\"><tbody><row><entry>A</entry></row></tbody></tgroup></table></section></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        OfficeDocumentModelTable table = Assert.Single(model.Tables);
        OfficeDocumentModelNode tableNode = FindStructureNode(model.Structure, "table");
        tableNode.Location.HeadingPath = table.Location!.HeadingPath;

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model, profile: profile);

        Assert.Single(converted.Value.Xml.Descendants(), element => element.Name.LocalName == "table");
        Assert.DoesNotContain(converted.Diagnostics, diagnostic => diagnostic.Code == "DB122");
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

        XElement author = converted.Xml.Root!.Elements().Single(element => element.Name.LocalName == "info")
            .Descendants().Single(element => element.Name.LocalName == "author");
        Assert.Equal("Jane Doe", author.Elements().Single(element => element.Name.LocalName == "personname").Value);
        Assert.DoesNotContain(author.Nodes().OfType<XText>(), text => !string.IsNullOrWhiteSpace(text.Value));
    }

    [Theory]
    [InlineData("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><info><authorgroup><author>Jane Doe</author></authorgroup></info></article>")]
    [InlineData("<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><articleinfo><authorgroup><author>Jane Doe</author></authorgroup></articleinfo></article>")]
    public void SharedRoundTripDoesNotDuplicateAuthorsNestedInRootMetadata(string source) {
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;

        DocBookDocument restored = DocBookDocument.FromOfficeDocumentModel(model).Value;

        Assert.Equal("Jane Doe", model.Source.Author);
        XElement author = Assert.Single(restored.Xml.Descendants(), element => element.Name.LocalName == "author");
        Assert.Equal("Jane Doe", Assert.Single(author.Elements(), element => element.Name.LocalName == "personname").Value);
        Assert.DoesNotContain(author.Nodes().OfType<XText>(), text => !string.IsNullOrWhiteSpace(text.Value));
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
        const string docBookFive = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:x=\"urn:extension\" version=\"5.2\"><info><author><personname xml:id=\"five-person\"><firstname>Jane</firstname><surname>Doe</surname></personname><affiliation><orgname>Example</orgname></affiliation></author></info><indexterm><primary>topic</primary></indexterm><custom flag=\"five\">native</custom><x:box><x:item>value</x:item></x:box></article>";
        DocBookDocument asFour = DocBookDocument.FromOfficeDocumentModel(
            DocBookDocument.Parse(docBookFive).ToOfficeDocumentModel().Value,
            profile: DocBookProfile.DocBook45).Value;
        XElement[] fourVocabulary = asFour.Xml.Descendants()
            .Where(element => new[] { "personname", "firstname", "surname", "affiliation", "orgname", "primary" }.Contains(element.Name.LocalName)).ToArray();
        Assert.Equal(6, fourVocabulary.Length);
        Assert.All(fourVocabulary,
            element => Assert.Equal(XNamespace.None, element.Name.Namespace));
        Assert.Equal("five-person", (string?)fourVocabulary.Single(element => element.Name.LocalName == "personname").Attribute("id"));
        Assert.Null(fourVocabulary.Single(element => element.Name.LocalName == "personname").Attribute(XNamespace.Xml + "id"));
        XElement fourCustom = asFour.Xml.Descendants().Single(element => element.Name.LocalName == "custom");
        Assert.Equal(DocBookSchemaProfiles.DocBook52.NamespaceUri, fourCustom.Name.NamespaceName);
        Assert.Equal("five", (string?)fourCustom.Attribute("flag"));
        Assert.Equal("urn:extension", asFour.Xml.Descendants().Single(element => element.Name.LocalName == "box").Name.NamespaceName);
        Assert.DoesNotContain(asFour.Validate().Diagnostics, diagnostic => diagnostic.Code == "DB023");

        const string docBookFour = "<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article><articleinfo><author><personname id=\"four-person\"><firstname>Jane</firstname><surname>Doe</surname></personname><affiliation><orgname>Example</orgname></affiliation></author></articleinfo><indexterm><primary>topic</primary></indexterm><custom flag=\"four\">native</custom></article>";
        DocBookDocument asFive = DocBookDocument.FromOfficeDocumentModel(
            DocBookDocument.Parse(docBookFour).ToOfficeDocumentModel().Value,
            profile: DocBookProfile.DocBook52).Value;
        XElement[] fiveVocabulary = asFive.Xml.Descendants()
            .Where(element => new[] { "personname", "firstname", "surname", "affiliation", "orgname", "primary" }.Contains(element.Name.LocalName)).ToArray();
        Assert.Equal(6, fiveVocabulary.Length);
        Assert.All(fiveVocabulary,
            element => Assert.Equal(DocBookSchemaProfiles.DocBook52.NamespaceUri, element.Name.NamespaceName));
        Assert.Equal("four-person", (string?)fiveVocabulary.Single(element => element.Name.LocalName == "personname").Attribute(XNamespace.Xml + "id"));
        Assert.Null(fiveVocabulary.Single(element => element.Name.LocalName == "personname").Attribute("id"));
        XElement fiveCustom = asFive.Xml.Descendants().Single(element => element.Name.LocalName == "custom");
        Assert.Equal(XNamespace.None, fiveCustom.Name.Namespace);
        Assert.Equal("four", (string?)fiveCustom.Attribute("flag"));
        Assert.DoesNotContain(asFive.Validate().Diagnostics, diagnostic => diagnostic.Code == "DB023");
    }

    [Theory]
    [InlineData("<!DOCTYPE article PUBLIC \"-//OASIS//DTD DocBook XML V4.5//EN\" \"http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd\"><article xmlns:d=\"http://docbook.org/ns/docbook\"><articleinfo><author><d:affiliation><d:orgname>Wrong</d:orgname></d:affiliation></author></articleinfo></article>")]
    [InlineData("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><info><author><affiliation xmlns=\"\"><orgname>Wrong</orgname></affiliation></author></info></article>")]
    public void ValidationRejectsKnownVocabularyFromTheOtherProfileNamespace(string source) {
        Assert.Contains(DocBookDocument.Parse(source).Validate().Diagnostics, diagnostic =>
            diagnostic.Code == "DB023" && diagnostic.Severity == DocBookDiagnosticSeverity.Error);
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

    [Theory]
    [InlineData(DocBookProfile.DocBook45, "sectioninfo")]
    [InlineData(DocBookProfile.DocBook52, "info")]
    public void TypedSectionMetadataPrecedesTitle(DocBookProfile profile, string expectedName) {
        DocBookDocument created = DocBookDocument.CreateArticle(profile);
        DocBookNode section = created.AddSection("Section");
        DocBookNode info = section.Add(DocBookNodeKind.Info);

        Assert.Equal(expectedName, info.Name);
        Assert.Equal(expectedName, section.Children.First().Name);
        Assert.Equal(DocBookNodeKind.Info, DocBookDocument.Parse(created.ToDocBook()).Root.Children
            .Single(node => node.Kind == DocBookNodeKind.Section).Children.First().Kind);
    }

    [Fact]
    public void ProfileConversionUsesSectionInfoForDocBookFour() {
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

    [Fact]
    public void SharedReverseConversionRejectsCyclesAndConfiguredStructureLimits() {
        var cyclic = new OfficeDocumentModelNode { Kind = "section", Text = "Cycle" };
        cyclic.Children = new[] { cyclic };
        var cyclicModel = new OfficeDocumentModel { Format = OfficeDocumentFormat.DocBook, Structure = new[] { cyclic } };
        Assert.Throws<InvalidDataException>(() => DocBookDocument.FromOfficeDocumentModel(cyclicModel));

        var root = new OfficeDocumentModelNode { Kind = "section", Text = "Root" };
        var child = new OfficeDocumentModelNode { Kind = "section", Text = "Child" };
        var grandchild = new OfficeDocumentModelNode { Kind = "paragraph", Text = "Grandchild" };
        root.Children = new[] { child };
        child.Children = new[] { grandchild };
        var deepModel = new OfficeDocumentModel { Format = OfficeDocumentFormat.DocBook, Structure = new[] { root } };

        Assert.Throws<InvalidDataException>(() => DocBookDocument.FromOfficeDocumentModel(
            deepModel, options: new DocBookConversionOptions { MaxStructureDepth = 2 }));
        Assert.Throws<InvalidDataException>(() => DocBookDocument.FromOfficeDocumentModel(
            deepModel, options: new DocBookConversionOptions { MaxStructureNodes = 2 }));
    }

    [Fact]
    public void SharedForwardConversionBoundsNativeEdits() {
        DocBookDocument deep = DocBookDocument.CreateArticle();
        XElement parent = deep.Xml.Root!;
        for (int depth = 0; depth < 8; depth++) {
            var child = new XElement(deep.Xml.Root!.Name.Namespace + "section");
            parent.Add(child);
            parent = child;
        }
        Assert.Throws<InvalidDataException>(() => deep.ToOfficeDocumentModel(options:
            new DocBookConversionOptions { MaxStructureDepth = 4 }));

        DocBookDocument metadataDeep = DocBookDocument.CreateArticle();
        XElement metadataParent = metadataDeep.Xml.Root!;
        for (int depth = 0; depth < 8; depth++) {
            var child = new XElement(metadataDeep.Xml.Root!.Name.Namespace + "info");
            metadataParent.Add(child);
            metadataParent = child;
        }
        Assert.Throws<InvalidDataException>(() => metadataDeep.ToOfficeDocumentModel(options:
            new DocBookConversionOptions { MaxStructureDepth = 4 }));

        DocBookDocument wide = DocBookDocument.CreateArticle();
        wide.Xml.Root!.Add(
            new XElement(wide.Xml.Root.Name.Namespace + "para", "One"),
            new XElement(wide.Xml.Root.Name.Namespace + "para", "Two"));
        Assert.Throws<InvalidDataException>(() => wide.ToOfficeDocumentModel(options:
            new DocBookConversionOptions { MaxStructureNodes = 2 }));
    }

    [Theory]
    [InlineData("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para>Old<emphasis> value</emphasis></para></article>", "paragraph", "para")]
    [InlineData("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><section><title>Old<emphasis> value</emphasis></title></section></article>", "title", "title")]
    [InlineData("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><programlisting>Old<emphasis> value</emphasis></programlisting></article>", "code", "programlisting")]
    [InlineData("<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><para><link xlink:href=\"https://example.test\">Old<emphasis> value</emphasis></link></para></article>", "link", "link")]
    [InlineData("<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><info><author><personname><firstname>Old</firstname><surname>Name</surname></personname></author></info></article>", "author", "author")]
    public void SharedReverseConversionReportsAndPreservesEditedPrimaryText(string source, string kind, string localName) {
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        FindStructureNode(model.Structure, kind).Text = "Edited";

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB125");
        XElement element = Assert.Single(converted.Value.Xml.Descendants(), candidate => candidate.Name.LocalName == localName);
        Assert.Equal("Edited", element.Value);
    }

    [Fact]
    public void SharedReverseConversionPreservesEditedStructuralTitlesAndTheirBodyContent() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><section><title>Section</title><para>Section body</para></section><table><title>Table</title><tgroup cols=\"1\"><tbody><row><entry>Cell</entry></row></tbody></tgroup></table><figure><title>Figure</title><mediaobject><imageobject><imagedata fileref=\"image.png\"/></imageobject></mediaobject></figure></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        FindStructureNode(model.Structure, "section").Text = "Edited section";
        FindStructureNode(model.Structure, "table").Text = "Edited table";
        FindStructureNode(model.Structure, "figure").Text = "Edited figure";

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.Equal("Edited section", converted.Value.Xml.Descendants().Single(element =>
            element.Name.LocalName == "section").Elements().Single(element => element.Name.LocalName == "title").Value);
        Assert.Equal("Section body", converted.Value.Xml.Descendants().Single(element => element.Name.LocalName == "para").Value);
        XElement table = Assert.Single(converted.Value.Xml.Descendants(), element => element.Name.LocalName == "table");
        Assert.Equal("Edited table", table.Elements().Single(element => element.Name.LocalName == "title").Value);
        Assert.Equal("Cell", table.Descendants().Single(element => element.Name.LocalName == "entry").Value);
        XElement figure = Assert.Single(converted.Value.Xml.Descendants(), element => element.Name.LocalName == "figure");
        Assert.Equal("Edited figure", figure.Elements().Single(element => element.Name.LocalName == "title").Value);
        Assert.Equal("image.png", (string?)figure.Descendants().Single(element =>
            element.Name.LocalName == "imagedata").Attribute("fileref"));
        Assert.Equal(3, converted.Diagnostics.Count(diagnostic => diagnostic.Code == "DB125" &&
            diagnostic.Message.IndexOf("Primary text", StringComparison.Ordinal) >= 0));
    }

    [Fact]
    public void SharedReverseConversionPreservesSynchronizedPrimaryAndFlatEdits() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><para>Old <emphasis>paragraph</emphasis></para><para><link xlink:href=\"https://example.test\">Old <emphasis>link</emphasis></link></para></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        OfficeDocumentModelNode paragraph = FindStructureNode(model.Structure, "paragraph");
        paragraph.Text = "Edited paragraph";
        model.Blocks.Single(block => block.Id == paragraph.Id).Text = paragraph.Text;
        OfficeDocumentModelNode link = FindStructureNode(model.Structure, "link");
        link.Text = "Edited link";
        Assert.Single(model.Links).Text = link.Text;

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.Equal("Edited paragraph", converted.Value.Xml.Root!.Elements().First(element =>
            element.Name.LocalName == "para").Value);
        Assert.Equal("Edited link", Assert.Single(converted.Value.Xml.Descendants(), element =>
            element.Name.LocalName == "link").Value);
        Assert.Equal(2, converted.Diagnostics.Count(diagnostic => diagnostic.Code == "DB125" &&
            diagnostic.Message.IndexOf("Primary text", StringComparison.Ordinal) >= 0));
    }

    [Fact]
    public void SharedReverseConversionReconcilesDocumentTitleProjectionEdits() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><info><title>Original</title></info><para>Body</para></article>";

        OfficeDocumentModel sourceEdited = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        sourceEdited.Source.Title = "Edited source";
        DocBookConversionResult<DocBookDocument> sourceResult = DocBookDocument.FromOfficeDocumentModel(sourceEdited);

        Assert.Equal("Edited source", sourceResult.Value.Title);
        Assert.Contains(sourceResult.Diagnostics, diagnostic => diagnostic.Code == "DB125" &&
            diagnostic.Message.IndexOf("Source.Title", StringComparison.Ordinal) >= 0);

        OfficeDocumentModel structureEdited = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        FindStructureNode(structureEdited.Structure, "title").Text = "Edited structure";
        DocBookConversionResult<DocBookDocument> structureResult = DocBookDocument.FromOfficeDocumentModel(structureEdited);

        Assert.Equal("Edited structure", structureResult.Value.Title);
        Assert.Contains(structureResult.Diagnostics, diagnostic => diagnostic.Code == "DB125" &&
            diagnostic.Message.IndexOf("recursive title", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void SharedReverseConversionReconcilesDocumentAuthorProjectionEdits() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><info><author><personname>Original</personname></author></info><para>Body</para></article>";

        OfficeDocumentModel sourceEdited = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        sourceEdited.Source.Author = "Edited source";
        DocBookConversionResult<DocBookDocument> sourceResult = DocBookDocument.FromOfficeDocumentModel(sourceEdited);

        Assert.Equal("Edited source", Assert.Single(sourceResult.Value.Xml.Descendants(), element =>
            element.Name.LocalName == "author").Value);
        Assert.Contains(sourceResult.Diagnostics, diagnostic => diagnostic.Code == "DB125" &&
            diagnostic.Message.IndexOf("Source.Author", StringComparison.Ordinal) >= 0);

        OfficeDocumentModel metadataEdited = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        metadataEdited.Metadata.Single(entry => entry.Category == "docbook" && entry.Name == "author").Value = "Edited metadata";
        DocBookConversionResult<DocBookDocument> metadataResult = DocBookDocument.FromOfficeDocumentModel(metadataEdited);

        Assert.Equal("Edited metadata", Assert.Single(metadataResult.Value.Xml.Descendants(), element =>
            element.Name.LocalName == "author").Value);
        Assert.Contains(metadataResult.Diagnostics, diagnostic => diagnostic.Code == "DB125");

        OfficeDocumentModel sourceRemoved = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        sourceRemoved.Source.Author = null;
        DocBookConversionResult<DocBookDocument> removalResult = DocBookDocument.FromOfficeDocumentModel(sourceRemoved);

        Assert.DoesNotContain(removalResult.Value.Xml.Descendants(), element => element.Name.LocalName == "author");
        Assert.Contains(removalResult.Diagnostics, diagnostic => diagnostic.Code == "DB125");
    }

    [Fact]
    public void SharedReverseConversionReportsInvalidKnownNodePlacement() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Structure = new[] { new OfficeDocumentModelNode { Kind = "list-item", Text = "Orphan" } }
        };

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.True(converted.HasLoss);
        Assert.Contains(converted.Diagnostics, diagnostic =>
            diagnostic.Code == "DB015" && diagnostic.Severity == DocBookDiagnosticSeverity.Error);
    }

    [Fact]
    public void SharedReverseConversionPreservesEditedFlatLinkKind() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para><xref linkend=\"target\"/></para></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        Assert.Single(model.Links).Kind = "link";

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB122" &&
            diagnostic.Message.IndexOf("link", StringComparison.OrdinalIgnoreCase) >= 0);
        Assert.Contains(converted.Value.Xml.Descendants(), element =>
            element.Name.LocalName == "link" && (string?)element.Attribute("linkend") == "target");
    }

    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void SharedReverseConversionEmitsFlatCrossReferencesAsXref(DocBookProfile profile) {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Links = new[] {
                new OfficeDocumentModelLink { Id = "xref", Kind = "cross-reference", DestinationName = "target" }
            }
        };

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(
            model, DocBookDocumentKind.Article, profile);
        XElement crossReference = Assert.Single(converted.Value.Xml.Descendants(), element =>
            element.Name.LocalName == "xref");

        Assert.Equal("target", (string?)crossReference.Attribute("linkend"));
        Assert.Empty(crossReference.Nodes());
        Assert.DoesNotContain(converted.Diagnostics, diagnostic => diagnostic.Code == "DB120");
        Assert.True(converted.Value.Validate().IsValid);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedReverseConversionSuppressesStaleFlatTargetsAfterRecursiveLinkEdits(bool internalTarget) {
        string source = internalTarget
            ? "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><para><link linkend=\"original\">Target</link></para></article>"
            : "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:xl=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><para><link xl:href=\"https://example.test/original\">Target</link></para></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        OfficeDocumentModelNode recursiveLink = model.Structure.Single().Children.Single(node => node.Kind == "link");
        var attributes = recursiveLink.Attributes.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal);
        string attributeName = internalTarget ? "linkend" : "{http://www.w3.org/1999/xlink}href";
        string editedTarget = internalTarget ? "edited" : "https://example.test/edited";
        attributes[attributeName] = editedTarget;
        recursiveLink.Attributes = attributes;

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);
        XElement link = Assert.Single(converted.Value.Xml.Descendants(), element => element.Name.LocalName == "link");

        Assert.Equal(editedTarget, (string?)link.Attribute(XName.Get(attributeName)));
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB125" &&
            diagnostic.Message.IndexOf("conflicting targets", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    [Fact]
    public void SharedReverseConversionDiagnosesUnrepresentableLinkGeometry() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" xmlns:xl=\"http://www.w3.org/1999/xlink\" version=\"5.2\"><para><link xl:href=\"https://example.test/\">Site</link></para></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        Assert.Single(model.Links).Region = new OfficeDocumentModelRegion { X = 1, Y = 2, Width = 3, Height = 4 };

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.True(converted.HasLoss);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB120" &&
            diagnostic.Message.IndexOf("could not be represented", StringComparison.OrdinalIgnoreCase) >= 0);
        Assert.Contains(converted.Value.Xml.Descendants(), element =>
            (string?)element.Attribute(XName.Get("href", "http://www.w3.org/1999/xlink")) == "https://example.test/");
    }

    [Fact]
    public void SharedBookConversionDoesNotCreateAChapterForAnOmittedFlatLink() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Metadata = new[] {
                new OfficeDocumentModelMetadataEntry { Category = "docbook", Name = "kind", Value = "book" }
            },
            Links = new[] {
                new OfficeDocumentModelLink { Id = "missing-target", Kind = "link", Text = "Unavailable" }
            }
        };

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.DoesNotContain(converted.Value.Xml.Root!.Elements(), element => element.Name.LocalName == "chapter");
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB120" &&
            diagnostic.Message.IndexOf("no DocBook-representable", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    [Fact]
    public void SharedReverseConversionPreservesLeafCaptionTextAsParagraphContent() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Structure = new[] {
                new OfficeDocumentModelNode {
                    Kind = "media",
                    Children = new OfficeDocumentModelNode[] {
                        new OfficeDocumentModelNode {
                            Kind = "image-object",
                            Children = new[] {
                                new OfficeDocumentModelNode {
                                    Kind = "image",
                                    Attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                                        ["fileref"] = "assets/chart.png"
                                    }
                                }
                            }
                        },
                        new OfficeDocumentModelNode { Kind = "caption", Text = "Visible caption" }
                    }
                }
            }
        };

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);
        XElement caption = Assert.Single(converted.Value.Xml.Descendants(), element => element.Name.LocalName == "caption");

        Assert.Equal("Visible caption", Assert.Single(caption.Elements(), element => element.Name.LocalName == "para").Value);
        Assert.DoesNotContain(converted.Diagnostics, diagnostic => diagnostic.Code == "DB116");
        Assert.True(converted.Value.Validate().IsValid);
    }

    [Fact]
    public void SharedConversionDoesNotReportFlatFallbackWithoutFlatContent() {
        var empty = new OfficeDocumentModel { Format = OfficeDocumentFormat.DocBook };
        var titleOnly = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Source = new OfficeDocumentModelSource { Title = "Guide" }
        };

        DocBookConversionResult<DocBookDocument> emptyResult = DocBookDocument.FromOfficeDocumentModel(empty);
        DocBookConversionResult<DocBookDocument> titleResult = DocBookDocument.FromOfficeDocumentModel(titleOnly);

        Assert.False(emptyResult.HasLoss);
        Assert.False(titleResult.HasLoss);
        Assert.Equal("Guide", titleResult.Value.Title);
        Assert.DoesNotContain(titleResult.Diagnostics, diagnostic => diagnostic.Code == "DB103");
    }

    [Fact]
    public void SharedConversionRepresentsAuthorMetadataAndDiagnosesUnsupportedMetadata() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Structure = new[] { new OfficeDocumentModelNode { Kind = "paragraph", Text = "Body" } },
            Metadata = new[] {
                new OfficeDocumentModelMetadataEntry { Category = "docbook", Name = "author", Value = "Jane Doe" },
                new OfficeDocumentModelMetadataEntry { Category = "portable", Name = "subject", Value = "Example" }
            }
        };

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.Equal("Jane Doe", converted.Value.Xml.Descendants().Single(element => element.Name.LocalName == "personname").Value);
        Assert.Single(converted.Diagnostics, diagnostic => diagnostic.Code == "DB125" &&
            diagnostic.Message.IndexOf("portable/subject", StringComparison.Ordinal) >= 0);
    }

    [Theory]
    [InlineData("4.5", DocBookProfile.DocBook52, "id", "{http://www.w3.org/XML/1998/namespace}id")]
    [InlineData("5.2", DocBookProfile.DocBook45, "{http://www.w3.org/XML/1998/namespace}id", "id")]
    public void SharedProfileConversionNormalizesIdentifierAttributes(
        string sourceProfile,
        DocBookProfile targetProfile,
        string sourceAttribute,
        string targetAttribute) {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Metadata = new[] { new OfficeDocumentModelMetadataEntry { Category = "docbook", Name = "profile", Value = sourceProfile } },
            Structure = new[] {
                new OfficeDocumentModelNode {
                    Kind = "section",
                    Text = "Section",
                    Attributes = new Dictionary<string, string> { [sourceAttribute] = "section-id" }
                }
            }
        };

        DocBookDocument converted = DocBookDocument.FromOfficeDocumentModel(model, profile: targetProfile).Value;
        XElement section = Assert.Single(converted.Xml.Descendants(), element => element.Name.LocalName == "section");

        Assert.Equal("section-id", section.Attribute(XName.Get(targetAttribute))?.Value);
        Assert.Null(section.Attribute(XName.Get(sourceAttribute)));
    }

    [Fact]
    public void SharedReverseConversionPreservesEditedFlatTableProjection() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup cols=\"1\"><tbody><row><entry>Original</entry></row></tbody></tgroup></informaltable></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        OfficeDocumentModelTable table = Assert.Single(model.Tables);
        table.Rows = new[] { new[] { "Edited" } };

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB122" &&
            diagnostic.Message.IndexOf("table", StringComparison.OrdinalIgnoreCase) >= 0);
        Assert.Contains(converted.Value.Xml.Descendants(), element => element.Name.LocalName == "entry" && element.Value == "Edited");
    }

    [Fact]
    public void SharedReverseConversionReportsEditedFlatTableSummary() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><informaltable><tgroup cols=\"1\"><tbody><row><entry>Value</entry></row></tbody></tgroup></informaltable></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        Assert.Single(model.Tables).Summary = "Edited summary";

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB122" &&
            diagnostic.Message.IndexOf("table", StringComparison.OrdinalIgnoreCase) >= 0);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB124" &&
            diagnostic.Message.IndexOf("summary", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedReverseConversionOmitsFlatTablesWithoutCalsBodyContent(bool includeEmptyRow) {
        var table = new OfficeDocumentModelTable {
            Columns = new[] { "Value" },
            Rows = includeEmptyRow
                ? new IReadOnlyList<string>[] { Array.Empty<string>() }
                : Array.Empty<IReadOnlyList<string>>()
        };
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Tables = new[] { table }
        };

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.True(converted.HasLoss);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB126");
        Assert.DoesNotContain(converted.Value.Xml.Descendants(), element =>
            element.Name.LocalName == "table" || element.Name.LocalName == "informaltable");
        Assert.True(converted.Value.Validate().IsValid);
    }

    [Fact]
    public void SharedReverseConversionPreservesEditedFlatAssetProjection() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><mediaobject><imageobject><imagedata fileref=\"assets/original.png\"/></imageobject><textobject><phrase>Original alt</phrase></textobject><caption>Original caption</caption></mediaobject></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        OfficeDocumentModelAsset asset = Assert.Single(model.Assets);
        asset.FileName = "edited.png";
        asset.Title = "Edited caption";
        asset.AltText = "Edited alt";

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB122" &&
            diagnostic.Message.IndexOf("asset", StringComparison.OrdinalIgnoreCase) >= 0);
        XElement image = Assert.Single(converted.Value.Xml.Descendants(), element =>
            element.Name.LocalName == "imagedata" && (string?)element.Attribute("fileref") == "edited.png");
        XElement media = image.Ancestors().Single(element => element.Name.LocalName == "mediaobject");
        Assert.Equal("Edited caption", media.Elements().Single(element => element.Name.LocalName == "caption").Value);
        Assert.Equal("Edited alt", media.Elements().Single(element => element.Name.LocalName == "textobject").Value);
    }

    [Fact]
    public void SharedReverseConversionPreservesSourceReferenceOnlyAssetEdits() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><mediaobject><imageobject><imagedata fileref=\"assets/original.png\"/></imageobject></mediaobject></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        Assert.Single(model.Assets).SourceObjectId = "assets/edited.jpg";

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB122" &&
            diagnostic.Message.IndexOf("asset", StringComparison.OrdinalIgnoreCase) >= 0);
        Assert.DoesNotContain(converted.Diagnostics, diagnostic => diagnostic.Code == "DB124");
        Assert.Contains(converted.Value.Xml.Descendants(), element =>
            element.Name.LocalName == "imagedata" && (string?)element.Attribute("fileref") == "assets/edited.jpg");
    }

    [Fact]
    public void SharedReverseConversionTreatsMissingDerivedImageMetadataAsUnspecified() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><mediaobject><imageobject><imagedata fileref=\"assets/original.png\"/></imageobject></mediaobject></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        OfficeDocumentModelAsset asset = Assert.Single(model.Assets);
        asset.Extension = null;
        asset.MediaType = null;

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.DoesNotContain(converted.Diagnostics, diagnostic => diagnostic.Code == "DB122" || diagnostic.Code == "DB124");
        Assert.Single(converted.Value.Xml.Descendants(), element =>
            element.Name.LocalName == "imagedata" && (string?)element.Attribute("fileref") == "assets/original.png");
    }

    [Fact]
    public void SharedReverseConversionDiagnosesConflictingAssetReferenceEdits() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><mediaobject><imageobject><imagedata fileref=\"assets/original.png\"/></imageobject></mediaobject></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        OfficeDocumentModelAsset asset = Assert.Single(model.Assets);
        asset.SourceObjectId = "assets/source-edit.png";
        asset.FileName = "filename-edit.png";

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB125" &&
            diagnostic.Message.IndexOf("SourceObjectId", StringComparison.Ordinal) >= 0);
        Assert.Contains(converted.Value.Xml.Descendants(), element =>
            element.Name.LocalName == "imagedata" && (string?)element.Attribute("fileref") == "assets/source-edit.png");
    }

    [Fact]
    public void SharedReverseConversionReportsEditedUnsupportedAssetFields() {
        const string source = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><mediaobject><imageobject><imagedata fileref=\"figure.png\"/></imageobject></mediaobject></article>";
        OfficeDocumentModel model = DocBookDocument.Parse(source).ToOfficeDocumentModel().Value;
        OfficeDocumentModelAsset asset = Assert.Single(model.Assets);
        asset.PayloadBytes = new byte[] { 1, 2, 3 };
        asset.Width = 640;

        DocBookConversionResult<DocBookDocument> converted = DocBookDocument.FromOfficeDocumentModel(model);

        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB122" &&
            diagnostic.Message.IndexOf("asset", StringComparison.OrdinalIgnoreCase) >= 0);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "DB124" &&
            diagnostic.Message.IndexOf("payload", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    [Fact]
    public void SharedReverseConversionReusesKindLookupForLargeStructures() {
        const int count = 25_000;
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Structure = Enumerable.Range(0, count).Select(index => new OfficeDocumentModelNode {
                Id = "paragraph-" + index, Kind = "paragraph", Text = "Paragraph " + index
            }).ToArray()
        };

        var stopwatch = Stopwatch.StartNew();
        DocBookDocument converted = DocBookDocument.FromOfficeDocumentModel(model).Value;
        stopwatch.Stop();

        Assert.Equal(count, converted.Xml.Descendants().Count(element => element.Name.LocalName == "para"));
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(10),
            $"Indexed reverse conversion took {stopwatch.Elapsed}.");
    }

    private static OfficeDocumentModelNode FindStructureNode(IEnumerable<OfficeDocumentModelNode> nodes, string kind) {
        OfficeDocumentModelNode? result = FindStructureNodeOrDefault(nodes, kind);
        return result ?? throw new InvalidOperationException($"Shared structure node '{kind}' was not found.");
    }

    private static OfficeDocumentModelNode? FindStructureNodeOrDefault(IEnumerable<OfficeDocumentModelNode> nodes, string kind) {
        foreach (OfficeDocumentModelNode node in nodes) {
            if (string.Equals(node.Kind, kind, StringComparison.OrdinalIgnoreCase)) return node;
            OfficeDocumentModelNode? descendant = FindStructureNodeOrDefault(node.Children, kind);
            if (descendant != null) return descendant;
        }
        return null;
    }

}
