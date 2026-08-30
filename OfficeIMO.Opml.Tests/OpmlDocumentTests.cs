using System.Collections.Generic;
using System.Diagnostics;

namespace OfficeIMO.Opml.Tests;

public sealed class OpmlDocumentTests {
    [Theory]
    [InlineData(OpmlVersion.Opml10, "1.0")]
    [InlineData(OpmlVersion.Opml20, "2.0")]
    public void CreatesBothWriterProfilesAndReadsDeclaredOnePointOne(OpmlVersion version, string declaration) {
        OpmlDocument created = OpmlDocument.Create(version);
        Assert.Equal(declaration, created.DeclaredVersion);
        Assert.Equal(version, OpmlDocument.Parse(created.ToOpml()).Version);

        OpmlDocument onePointOne = OpmlDocument.Parse("<opml version=\"1.1\"><head/><body/></opml>");
        Assert.Equal("1.1", onePointOne.DeclaredVersion);
        Assert.Equal(OpmlVersion.Opml10, onePointOne.Version);
        Assert.Contains(onePointOne.Validate().Diagnostics, diagnostic => diagnostic.Code == "OPML002");
    }

    [Fact]
    public void CreateEditValidateAndReopenNestedSubscriptionList() {
        OpmlDocument document = OpmlDocument.Create();
        document.Head.Title = "Subscriptions";
        OpmlOutline folder = document.AddOutline("Technology");
        OpmlOutline feed = folder.AddChild("OfficeIMO");
        feed.Type = "rss"; feed.XmlUrl = "https://example.test/feed.xml"; feed.HtmlUrl = "https://example.test/";

        Assert.True(document.Validate().IsValid);
        OpmlDocument reopened = OpmlDocument.Parse(document.ToOpml());
        Assert.Equal("Subscriptions", reopened.Head.Title);
        Assert.Equal("https://example.test/feed.xml", reopened.Outlines.Single().Children.Single().XmlUrl);
    }

    [Theory]
    [InlineData(OpmlVersion.Opml10)]
    [InlineData(OpmlVersion.Opml20)]
    public void TypedHeadSettersMaintainStandardSchemaOrder(OpmlVersion version) {
        OpmlDocument document = OpmlDocument.Create(version);
        XElement head = document.Xml.Root!.Element("head")!;
        head.Add(new XElement("ownerId", "https://example.test/owner"));
        document.Head.ExpansionState = "1";
        document.Head.Docs = "https://example.test/docs";
        document.Head.OwnerEmail = "owner@example.test";
        document.Head.OwnerName = "Owner";
        document.Head.DateModified = "modified";
        document.Head.DateCreated = "created";
        document.Head.Title = "Title";
        document.AddOutline("Outline");

        Assert.Equal(new[] {
            "title", "dateCreated", "dateModified", "ownerName", "ownerEmail", "ownerId", "docs", "expansionState"
        }, head.Elements().Select(element => element.Name.LocalName));
        Assert.True(document.Validate().IsValid);
    }

    [Theory]
    [InlineData("1.0", "<ownerName>Owner</ownerName><title>Title</title>", "OPML006")]
    [InlineData("2.0", "<ownerName>Owner</ownerName><title>Title</title>", "OPML006")]
    [InlineData("1.0", "<ownerName>One</ownerName><ownerName>Two</ownerName>", "OPML007")]
    [InlineData("2.0", "<ownerName>One</ownerName><ownerName>Two</ownerName>", "OPML007")]
    public void ValidationRejectsMisorderedOrDuplicateStandardHeadFields(
        string version,
        string headContent,
        string expectedCode) {
        OpmlDocument document = OpmlDocument.Parse(
            $"<opml version=\"{version}\"><head>{headContent}</head><body><outline text=\"Outline\"/></body></opml>");

        Assert.Contains(document.Validate().Diagnostics, diagnostic =>
            diagnostic.Code == expectedCode && diagnostic.Severity == OpmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void UnchangedInputIsExactAndEditPreservesUnknownContent() {
        const string source = "<?xml version=\"1.0\"?><opml version=\"2.0\" xmlns:x=\"urn:test\"><head><title>T</title><x:extra p=\"1\" /></head><body><!--keep--><outline text=\"A\" x:flag=\"yes\"><x:child /></outline></body></opml>";
        OpmlDocument document = OpmlDocument.Parse(source);
        Assert.Equal(source, document.ToOpml());

        document.Outlines.Single().Text = "B";
        string edited = document.ToOpml();
        Assert.Contains("x:flag=\"yes\"", edited);
        Assert.Contains("<x:child", edited);
        Assert.Contains("<!--keep-->", edited);
        Assert.Equal("1", document.Xml.Root!.Element("head")!.Elements().Single(e => e.Name.LocalName == "extra").Attribute("p")!.Value);
    }

    [Fact]
    public void SubscriptionAndLinkRequirementsAreValidated() {
        OpmlDocument document = OpmlDocument.Parse("<opml version=\"2.0\"><head/><body><outline text=\"Feed\" type=\"rss\"/><outline text=\"Link\" type=\"link\"/></body></opml>");
        OpmlValidationResult result = document.Validate();
        Assert.False(result.IsValid);
        Assert.Contains(result.Diagnostics, d => d.Code == "OPML011");
        Assert.Contains(result.Diagnostics, d => d.Code == "OPML012");
    }

    [Theory]
    [InlineData("<opml version=\"1.0\"><head/><body/></opml>")]
    [InlineData("<opml version=\"2.0\"><head/><body><extension><outline text=\"Nested\"/></extension></body></opml>")]
    public void ValidationRequiresADirectOutlineInTheBody(string source) {
        OpmlValidationResult validation = OpmlDocument.Parse(source).Validate();

        Assert.Contains(validation.Diagnostics, diagnostic => diagnostic.Code == "OPML013" &&
            diagnostic.Severity == OpmlDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData("<opml version=\"2.0\">root<head/><body><outline text=\"A\"/></body></opml>", "/opml")]
    [InlineData("<opml version=\"2.0\"><head>head</head><body><outline text=\"A\"/></body></opml>", "/opml/head")]
    [InlineData("<opml version=\"2.0\"><head/><body>body<outline text=\"A\"/></body></opml>", "/opml/body")]
    [InlineData("<opml version=\"2.0\"><head/><body><outline text=\"A\">outline</outline></body></opml>", "/opml/body//outline[1]")]
    public void ValidationRejectsDirectTextInElementOnlyContainers(string source, string path) {
        OpmlValidationResult validation = OpmlDocument.Parse(source).Validate();

        Assert.Contains(validation.Diagnostics, diagnostic => diagnostic.Code == "OPML014" &&
            diagnostic.Severity == OpmlDiagnosticSeverity.Error && diagnostic.Path == path);
    }

    [Fact]
    public void LimitsAndDtdPolicyRejectHostileInputs() {
        Assert.Throws<InvalidDataException>(() => OpmlDocument.Parse(
            "<opml version=\"2.0\"><head/><body><outline text=\"a\"/><outline text=\"b\"/></body></opml>",
            new OpmlReadOptions { MaxOutlines = 1 }));
        Assert.Throws<InvalidDataException>(() => OpmlDocument.Parse(
            "<opml version=\"2.0\"><head/><body><extension/></body></opml>",
            new OpmlReadOptions { MaxElements = 3 }));
        Assert.ThrowsAny<Exception>(() => OpmlDocument.Parse(
            "<!DOCTYPE opml [<!ENTITY x \"boom\">]><opml version=\"2.0\"><head/><body><outline text=\"&x;\"/></body></opml>"));
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => OpmlDocument.Parse("<opml version=\"2.0\"><head/><body/></opml>", cancellationToken: canceled.Token));
    }

    [Fact]
    public void MaxOutlinesStopsParsingBeforeTheRemainingXmlIsMaterialized() {
        const string source = "<opml version=\"2.0\"><head/><body><outline text=\"a\"/><outline text=\"b\"/><";
        var options = new OpmlReadOptions { MaxOutlines = 1 };

        InvalidDataException textException = Assert.Throws<InvalidDataException>(() => OpmlDocument.Parse(source, options));
        Assert.Contains("MaxOutlines", textException.Message, StringComparison.Ordinal);

        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(source));
        InvalidDataException streamException = Assert.Throws<InvalidDataException>(() => OpmlDocument.Load(stream, options));
        Assert.Contains("MaxOutlines", streamException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async System.Threading.Tasks.Task StreamLoadPreservesPositionAndAsyncWriteRewinds() {
        byte[] bytes = Encoding.UTF8.GetBytes("<opml version=\"2.0\"><head/><body><outline text=\"A\"/></body></opml>");
        using var input = new MemoryStream(bytes); input.Position = 5;
        OpmlDocument document = await OpmlDocument.LoadAsync(input);
        Assert.Equal(5, input.Position);
        using var output = new MemoryStream(new byte[512], writable: true);
        await document.WriteAsync(output);
        Assert.Equal(0, output.Position);
        Assert.Equal("A", OpmlDocument.Load(output).Outlines.Single().Text);
    }

    [Fact]
    public void LoadedBytesHonorDeclaredEncodingAndParsedTextWritesConsistentUtf8() {
        const string latinSource = "<?xml version=\"1.0\" encoding=\"iso-8859-1\"?><opml version=\"2.0\"><head><title>Caf\u00e9</title></head><body/></opml>";
        byte[] latinBytes = Encoding.GetEncoding("iso-8859-1").GetBytes(latinSource);
        OpmlDocument loaded = OpmlDocument.Load(new MemoryStream(latinBytes));
        Assert.Equal(latinSource, loaded.ToOpml());
        using var exact = new MemoryStream();
        loaded.Write(exact);
        Assert.Equal(latinBytes, exact.ToArray());

        const string utf16Declaration = "<?xml version=\"1.0\" encoding=\"utf-16\"?><opml version=\"2.0\"><head/><body/></opml>";
        OpmlDocument parsed = OpmlDocument.Parse(utf16Declaration);
        using var output = new MemoryStream();
        parsed.Write(output);
        string serialized = Encoding.UTF8.GetString(output.ToArray());
        Assert.Contains("encoding=\"utf-8\"", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(OpmlVersion.Opml20, OpmlDocument.Load(new MemoryStream(output.ToArray())).Version);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void LoadedBomlessUnicodeSourceIsReturnedAsTextWithoutUtf8Fallback(bool utf32, bool bigEndian) {
        const string source = "<opml version=\"2.0\"><head><title>Caf\u00e9</title></head><body/></opml>";
        Encoding encoding = utf32
            ? new UTF32Encoding(bigEndian, false, true)
            : new UnicodeEncoding(bigEndian, false, true);

        OpmlDocument document = OpmlDocument.Load(new MemoryStream(encoding.GetBytes(source)));

        Assert.Equal(source, document.ToOpml());
    }

    [Fact]
    public async System.Threading.Tasks.Task PathSaveAndAsyncLoadReopenTheCommittedArtifact() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-opml-" + Guid.NewGuid().ToString("N") + ".opml");
        try {
            OpmlDocument document = OpmlDocument.Create(); document.AddOutline("Path");
            await document.SaveAsync(path);
            using var stream = File.OpenRead(path);
            Assert.Equal("Path", (await OpmlDocument.LoadAsync(stream)).Outlines.Single().Text);
            Assert.Equal("Path", OpmlDocument.Load(path).Outlines.Single().Text);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void SharedModelRoundTripRetainsNestingAndAttributes() {
        OpmlDocument source = OpmlDocument.Create(OpmlVersion.Opml10);
        source.Head.Title = "Title";
        source.Head.OwnerName = "Jane Doe";
        source.Xml.Root!.Element("head")!.Element("title")!.SetAttributeValue(XName.Get("flag", "urn:test"), "metadata");
        source.Xml.Root!.Element("head")!.Element("ownerName")!.SetAttributeValue(XName.Get("flag", "urn:test"), "owner");
        OpmlOutline root = source.AddOutline("Root"); root.SetAttribute(XName.Get("flag", "urn:test"), "yes");
        root.Url = "https://example.test/root";
        root.AddChild("Child");
        var model = source.ToOfficeDocumentModel().Value;
        Assert.Equal(OfficeDocumentFormat.Opml, model.Format);
        Assert.Equal("Jane Doe", model.Source.Author);
        Assert.Equal("Child", model.Structure.Single().Children.Single().Text);
        Assert.Equal(2, model.Structure.SelectMany(node => new[] { node }.Concat(node.Children)).Select(node => node.Id).Distinct().Count());

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);
        Assert.False(converted.HasLoss);
        Assert.Equal(OpmlVersion.Opml10, converted.Value.Version);
        Assert.Equal("yes", converted.Value.Outlines.Single().GetAttribute(XName.Get("flag", "urn:test")));
        Assert.Equal("https://example.test/root", converted.Value.Outlines.Single().Url);
        Assert.DoesNotContain(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML107");
        Assert.Equal("metadata", converted.Value.Xml.Root!.Element("head")!.Element("title")!.Attribute(XName.Get("flag", "urn:test"))!.Value);
        Assert.Equal("Jane Doe", converted.Value.Head.OwnerName);
        Assert.Equal("owner", converted.Value.Xml.Root!.Element("head")!.Element("ownerName")!.Attribute(XName.Get("flag", "urn:test"))!.Value);
        Assert.Single(converted.Value.Xml.Root!.Element("head")!.Elements("ownerName"));
    }

    [Fact]
    public void SharedConversionMapsSourceAuthorToOwnerName() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Opml,
            Source = new OfficeDocumentModelSource { Author = "Alex Smith" }
        };

        OpmlDocument converted = OpmlDocument.FromOfficeDocumentModel(model).Value;

        Assert.Equal("Alex Smith", converted.Head.OwnerName);
    }

    [Theory]
    [InlineData("title", "Source title", "Metadata title")]
    [InlineData("ownerName", "Source owner", "Metadata owner")]
    public void SharedConversionDiagnosesConflictingPrimaryHeadProjections(
        string metadataName,
        string sourceValue,
        string metadataValue) {
        OpmlDocument source = OpmlDocument.Create();
        source.Head.Title = "Original title";
        source.Head.OwnerName = "Original owner";
        OfficeDocumentModel model = source.ToOfficeDocumentModel().Value;
        OfficeDocumentModelMetadataEntry metadata = model.Metadata.Single(entry =>
            entry.Category == "opml.head" && entry.Name == metadataName);
        metadata.Value = metadataValue;
        if (metadataName == "title") model.Source.Title = sourceValue;
        else model.Source.Author = sourceValue;

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML110" &&
            diagnostic.Message.IndexOf(metadataName, StringComparison.Ordinal) >= 0);
        Assert.Equal(sourceValue, metadataName == "title" ? converted.Value.Head.Title : converted.Value.Head.OwnerName);
    }

    [Theory]
    [InlineData("title")]
    [InlineData("ownerName")]
    public void SharedConversionPreservesClearedPrimaryHeadFields(string metadataName) {
        OpmlDocument source = OpmlDocument.Create();
        source.Head.Title = "Original title";
        source.Head.OwnerName = "Original owner";
        OfficeDocumentModel model = source.ToOfficeDocumentModel().Value;
        if (metadataName == "title") model.Source.Title = null;
        else model.Source.Author = null;

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        Assert.Null(metadataName == "title" ? converted.Value.Head.Title : converted.Value.Head.OwnerName);
        Assert.Null(converted.Value.Xml.Root!.Element("head")!.Element(metadataName));
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML110" &&
            diagnostic.Message.IndexOf(metadataName, StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void SharedConversionReportsUnsupportedVersionNormalization() {
        OpmlConversionResult<OfficeDocumentModel> forward = OpmlDocument.Parse(
            "<opml version=\"9.0\"><head/><body/></opml>").ToOfficeDocumentModel();

        OpmlConversionResult<OpmlDocument> restored = OpmlDocument.FromOfficeDocumentModel(forward.Value);

        Assert.True(restored.HasLoss);
        Assert.Contains(restored.Diagnostics, diagnostic => diagnostic.Code == "OPML105");
        Assert.Equal("2.0", restored.Value.DeclaredVersion);
    }

    [Fact]
    public void ParsingRejectsBodyBeforeHead() {
        Assert.Throws<InvalidDataException>(() =>
            OpmlDocument.Parse("<opml version=\"2.0\"><body/><head/></opml>"));
    }

    [Fact]
    public void ValidationReportsRootShapeProblemsAfterAdvancedEdits() {
        OpmlDocument renamed = OpmlDocument.Create();
        renamed.Xml.Root!.Name = "renamed";
        Assert.Contains(renamed.Validate().Diagnostics, diagnostic => diagnostic.Code == "OPML003");

        OpmlDocument missingBody = OpmlDocument.Create();
        missingBody.Xml.Root!.Element("body")!.Remove();
        OpmlValidationResult missingBodyResult = missingBody.Validate();
        Assert.False(missingBodyResult.IsValid);
        Assert.Contains(missingBodyResult.Diagnostics, diagnostic => diagnostic.Code == "OPML004");

        OpmlDocument reordered = OpmlDocument.Create();
        XElement head = reordered.Xml.Root!.Element("head")!;
        head.Remove();
        reordered.Xml.Root!.Add(head);
        Assert.Contains(reordered.Validate().Diagnostics, diagnostic => diagnostic.Code == "OPML005");
    }

    [Fact]
    public void SharedConversionBoundsCumulativeOutlineHeadingPaths() {
        OpmlDocument document = OpmlDocument.Create();
        OpmlOutline parent = document.AddOutline(new string('a', 2_000));
        parent.AddChild(new string('b', 2_000)).AddChild(new string('c', 2_000));

        OfficeDocumentModel model = document.ToOfficeDocumentModel().Value;
        OfficeDocumentModelNode[] nodes = model.Structure
            .SelectMany(root => new[] { root, root.Children[0], root.Children[0].Children[0] })
            .ToArray();

        Assert.All(nodes, node => Assert.True(node.Location.HeadingPath!.Length <= 1_024));
        Assert.All(model.Blocks, block => Assert.True(block.Location.HeadingPath!.Length <= 1_024));
    }

    [Fact]
    public void AdvancedXmlMutationCannotBeHiddenByUnchangedSourceFastPath() {
        const string source = "<opml version=\"2.0\"><head/><body><outline text=\"Before\"/></body></opml>";
        OpmlDocument document = OpmlDocument.Parse(source);
        document.Xml.Root!.Element("body")!.Element("outline")!.SetAttributeValue("text", "After");
        Assert.True(document.IsModified);
        Assert.Contains("text=\"After\"", document.ToOpml());
    }

    [Fact]
    public void SharedConversionReportsNativeOnlyExtensionContent() {
        OpmlDocument document = OpmlDocument.Parse("<opml version=\"2.0\" custom=\"x\" xmlns:x=\"urn:test\">root-text<head x:flag=\"head\">head-text</head><body x:flag=\"body\">body-text<!--native--><outline text=\"A\">outline-text<extra/></outline><body-extension/></body><root-extension/></opml>");
        OpmlConversionResult<OfficeDocumentModel> result = document.ToOfficeDocumentModel();
        Assert.True(result.HasLoss);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML200");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML202");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML203");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML204");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML205");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML206");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML207");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML208" && diagnostic.Path == "/opml/head");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML209" && diagnostic.Path == "/opml/body");
    }

    [Fact]
    public void RepeatedValidationAndConversionDiagnosticsAreBounded() {
        const string source = "<opml version=\"2.0\"><head/><body><outline text=\"1\" type=\"rss\"><extension/></outline><outline text=\"2\" type=\"rss\"><extension/></outline><outline text=\"3\" type=\"rss\"><extension/></outline><outline text=\"4\" type=\"rss\"><extension/></outline><outline text=\"5\" type=\"rss\"><extension/></outline></body></opml>";
        OpmlDocument document = OpmlDocument.Parse(source);

        OpmlValidationResult validation = document.Validate(new OpmlValidationOptions { MaxDetailedDiagnosticsPerCode = 2 });
        OpmlConversionResult<OfficeDocumentModel> conversion = document.ToOfficeDocumentModel(
            null, new OpmlConversionOptions { MaxDetailedDiagnosticsPerCode = 2 });
        var model = new OfficeDocumentModel {
            Structure = Enumerable.Range(1, 5)
                .Select(index => new OfficeDocumentModelNode { Kind = "unsupported-" + index, Text = "Value" }).ToArray()
        };
        OpmlConversionResult<OpmlDocument> reverse = OpmlDocument.FromOfficeDocumentModel(
            model, null, new OpmlConversionOptions { MaxDetailedDiagnosticsPerCode = 2 });

        Assert.Equal(3, validation.Diagnostics.Count(diagnostic => diagnostic.Code == "OPML011"));
        Assert.Contains(validation.Diagnostics, diagnostic => diagnostic.Code == "OPML011" &&
            diagnostic.Message.StartsWith("3 additional", StringComparison.Ordinal));
        Assert.Equal(3, conversion.Diagnostics.Count(diagnostic => diagnostic.Code == "OPML200"));
        Assert.Contains(conversion.Diagnostics, diagnostic => diagnostic.Code == "OPML200" &&
            diagnostic.Message.StartsWith("3 additional", StringComparison.Ordinal));
        Assert.Equal(3, reverse.Diagnostics.Count(diagnostic => diagnostic.Code == "OPML104"));
        Assert.Contains(reverse.Diagnostics, diagnostic => diagnostic.Code == "OPML104" &&
            diagnostic.Message.StartsWith("3 additional", StringComparison.Ordinal));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.Validate(
            new OpmlValidationOptions { MaxDetailedDiagnosticsPerCode = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.ToOfficeDocumentModel(
            null, new OpmlConversionOptions { MaxDetailedDiagnosticsPerCode = 0 }));
    }

    [Fact]
    public void SharedConversionReportsNonOutlineKindNormalization() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Structure = new[] { new OfficeDocumentModelNode { Kind = "paragraph", Text = "Body" } }
        };

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        Assert.True(converted.HasLoss);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML104");
        Assert.Equal("Body", converted.Value.Outlines.Single().Text);
    }

    [Fact]
    public void SharedModelPublishesSubscriptionAndOutlineLinks() {
        OpmlDocument document = OpmlDocument.Create();
        OpmlOutline outline = document.AddOutline("Feed");
        outline.Url = "https://example.test/direct";
        outline.XmlUrl = "https://example.test/feed.xml";
        outline.HtmlUrl = "https://example.test/";

        OfficeDocumentModel model = document.ToOfficeDocumentModel().Value;

        Assert.Equal(3, model.Links.Count);
        Assert.Contains(model.Links, link => link.Kind == "url" && link.Uri == outline.Url);
        Assert.Contains(model.Links, link => link.Kind == "subscription" && link.Uri == outline.XmlUrl);
        Assert.Contains(model.Links, link => link.Kind == "html" && link.Uri == outline.HtmlUrl);
    }

    [Fact]
    public void SharedConversionEmitsFlatLinksAsTypedOutlinesAndDiagnosesUnsupportedTargets() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Opml,
            Links = new[] {
                new OfficeDocumentModelLink { Id = "direct", Kind = "url", Text = "Direct", Uri = "https://example.test/direct" },
                new OfficeDocumentModelLink { Id = "feed", Kind = "subscription", Text = "Feed", Uri = "https://example.test/feed.xml" },
                new OfficeDocumentModelLink { Id = "site", Kind = "html", Text = "Site", Uri = "https://example.test/" },
                new OfficeDocumentModelLink { Id = "internal", Kind = "cross-reference", Text = "Target", DestinationName = "target" }
            }
        };

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        Assert.Contains(converted.Value.Outlines, outline => outline.Url == "https://example.test/direct");
        Assert.Contains(converted.Value.Outlines, outline => outline.XmlUrl == "https://example.test/feed.xml" && outline.Type == "rss");
        Assert.Contains(converted.Value.Outlines, outline => outline.HtmlUrl == "https://example.test/");
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML106" && diagnostic.Message.IndexOf("internal", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void SharedConversionDiagnosesUnrepresentableLinkGeometry() {
        OpmlDocument document = OpmlDocument.Create();
        document.AddOutline("Site").Url = "https://example.test/";
        OfficeDocumentModel model = document.ToOfficeDocumentModel().Value;
        Assert.Single(model.Links).Region = new OfficeDocumentModelRegion { X = 1, Y = 2, Width = 3, Height = 4 };

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        Assert.True(converted.HasLoss);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML106" &&
            diagnostic.Message.IndexOf("could not be represented", StringComparison.OrdinalIgnoreCase) >= 0);
        Assert.Contains(converted.Value.Outlines, outline => outline.Url == "https://example.test/");
    }

    [Fact]
    public void SharedConversionAppendsIndependentBlocksAndLinksAlongsideStructure() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Opml,
            Structure = new[] { new OfficeDocumentModelNode { Id = "outline-0", Kind = "outline", Text = "Structured" } },
            Blocks = new[] { new OfficeDocumentModelBlock { Id = "supplemental-block", Kind = "outline", Text = "Supplemental" } },
            Links = new[] {
                new OfficeDocumentModelLink { Id = "supplemental-link", Kind = "url", Text = "Site", Uri = "https://example.test/" }
            }
        };

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        Assert.Equal(new[] { "Structured", "Supplemental", "Site" }, converted.Value.Outlines.Select(outline => outline.Text));
        Assert.Equal("https://example.test/", converted.Value.Outlines[2].Url);
        Assert.Equal(2, converted.Diagnostics.Count(diagnostic => diagnostic.Code == "OPML107"));
    }

    [Fact]
    public void SharedConversionPreservesGeneratedLookingFlatContentWithoutProjectionProvenance() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Opml,
            Blocks = new[] {
                new OfficeDocumentModelBlock { Id = "outline-7", Kind = "outline", Text = "Independent block" }
            },
            Links = new[] {
                new OfficeDocumentModelLink {
                    Id = "opml-link-8-url", Kind = "url", Text = "Independent link", Uri = "https://example.test/"
                }
            }
        };

        OpmlDocument converted = OpmlDocument.FromOfficeDocumentModel(model).Value;

        Assert.Equal(new[] { "Independent block", "Independent link" }, converted.Outlines.Select(outline => outline.Text));
    }

    [Fact]
    public void SharedReverseConversionIndexesLargeDerivedFlatProjections() {
        const int count = 25_000;
        OfficeDocumentModelNode[] nodes = Enumerable.Range(0, count).Select(index => {
            string id = "outline-" + index;
            string text = "Outline " + index;
            return new OfficeDocumentModelNode {
                Id = id,
                Kind = "outline",
                Text = text,
                Attributes = new Dictionary<string, string> { ["url"] = "https://example.test/" + index }
            };
        }).ToArray();
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Opml,
            Structure = nodes,
            Blocks = nodes.Select(node => new OfficeDocumentModelBlock {
                Id = node.Id, Kind = node.Kind, Text = node.Text, Level = node.Level
            }).ToArray(),
            Links = nodes.Select((node, index) => new OfficeDocumentModelLink {
                Id = "opml-link-" + index + "-url", Kind = "url", Text = node.Text,
                Uri = node.Attributes["url"]
            }).ToArray()
        };

        var stopwatch = Stopwatch.StartNew();
        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);
        stopwatch.Stop();

        Assert.Equal(count, converted.Value.Outlines.Count);
        Assert.DoesNotContain(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML107");
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(10),
            $"Indexed reverse conversion took {stopwatch.Elapsed}.");
    }

    [Fact]
    public void SharedConversionDiagnosesUnsupportedTablesAndAssetsWithOrWithoutStructure() {
        var table = new OfficeDocumentModelTable { Title = "Values" };
        var asset = new OfficeDocumentModelAsset { Id = "figure", Kind = "image", FileName = "figure.png" };
        var structured = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Opml,
            Structure = new[] { new OfficeDocumentModelNode { Kind = "outline", Text = "Root" } },
            Tables = new[] { table },
            Assets = new[] { asset }
        };
        var flat = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Opml,
            Tables = new[] { table },
            Assets = new[] { asset }
        };

        OpmlConversionResult<OpmlDocument> structuredResult = OpmlDocument.FromOfficeDocumentModel(structured);
        OpmlConversionResult<OpmlDocument> flatResult = OpmlDocument.FromOfficeDocumentModel(flat);

        Assert.True(structuredResult.HasLoss);
        Assert.Equal(2, structuredResult.Diagnostics.Count(diagnostic => diagnostic.Code == "OPML108"));
        Assert.Equal(2, flatResult.Diagnostics.Count(diagnostic => diagnostic.Code == "OPML108"));
    }

    [Fact]
    public void SharedConversionDiagnosesUnsupportedPagesFormsAndVisuals() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Opml,
            Structure = new[] { new OfficeDocumentModelNode { Kind = "outline", Text = "Root" } },
            Pages = new[] { new OfficeDocumentModelPage { Number = 1, Name = "Page" } },
            Forms = new[] { new OfficeDocumentModelFormField { Id = "field", Kind = "text" } },
            Visuals = new[] { new OfficeDocumentModelVisual { Kind = "diagram", SourceName = "Visual" } }
        };

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        Assert.True(converted.HasLoss);
        Assert.Equal(3, converted.Diagnostics.Count(diagnostic => diagnostic.Code == "OPML108"));
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Message.IndexOf("page", StringComparison.OrdinalIgnoreCase) >= 0);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Message.IndexOf("form field", StringComparison.OrdinalIgnoreCase) >= 0);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Message.IndexOf("visual", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    [Fact]
    public void SharedConversionDiagnosesPortableOnlyBodyChannels() {
        var model = new OfficeDocumentModel { Markdown = "# Heading", Html = "<h1>Heading</h1>" };

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        Assert.True(converted.HasLoss);
        Assert.Empty(converted.Value.Outlines);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML108" &&
            diagnostic.Message.IndexOf("Markdown", StringComparison.Ordinal) >= 0);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML108" &&
            diagnostic.Message.IndexOf("HTML", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void SharedConversionDiagnosesUnsupportedSourceSubjectAndKeywords() {
        var model = new OfficeDocumentModel {
            Source = new OfficeDocumentModelSource { Subject = "Quarterly report", Keywords = "alpha, beta" }
        };

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        Assert.True(converted.HasLoss);
        Assert.Empty(converted.Value.Outlines);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML109" &&
            diagnostic.Message.IndexOf("Source.Subject", StringComparison.Ordinal) >= 0);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML109" &&
            diagnostic.Message.IndexOf("Source.Keywords", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void SharedConversionReportsInvalidEmptyBodyWithoutFlatFallbackForMetadataOnlyModels() {
        var empty = new OfficeDocumentModel { Format = OfficeDocumentFormat.Opml };
        var metadataOnly = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Opml,
            Source = new OfficeDocumentModelSource { Title = "Feeds", Author = "Owner" }
        };

        OpmlConversionResult<OpmlDocument> emptyResult = OpmlDocument.FromOfficeDocumentModel(empty);
        OpmlConversionResult<OpmlDocument> metadataResult = OpmlDocument.FromOfficeDocumentModel(metadataOnly);

        Assert.True(emptyResult.HasLoss);
        Assert.True(metadataResult.HasLoss);
        Assert.Equal("Feeds", metadataResult.Value.Head.Title);
        Assert.Equal("Owner", metadataResult.Value.Head.OwnerName);
        Assert.DoesNotContain(metadataResult.Diagnostics, diagnostic => diagnostic.Code == "OPML101");
        Assert.Contains(emptyResult.Diagnostics, diagnostic => diagnostic.Code == "OPML013");
        Assert.Contains(metadataResult.Diagnostics, diagnostic => diagnostic.Code == "OPML013");
    }

    [Fact]
    public void SemanticWalksObserveCancellation() {
        OpmlDocument document = OpmlDocument.Parse(
            "<opml version=\"2.0\"><head/><body><outline text=\"Root\"><outline text=\"Child\"/></outline></body></opml>");
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => document.Validate(null, cancellation.Token));
        Assert.Throws<OperationCanceledException>(() => document.ToOfficeDocumentModel(null, null, cancellation.Token));
    }

    [Fact]
    public void SharedReverseConversionRejectsCyclesAndConfiguredStructureLimits() {
        var cyclic = new OfficeDocumentModelNode { Kind = "outline", Text = "Cycle" };
        cyclic.Children = new[] { cyclic };
        var cyclicModel = new OfficeDocumentModel { Format = OfficeDocumentFormat.Opml, Structure = new[] { cyclic } };
        Assert.Throws<InvalidDataException>(() => OpmlDocument.FromOfficeDocumentModel(cyclicModel));

        var root = new OfficeDocumentModelNode { Kind = "outline", Text = "Root" };
        var child = new OfficeDocumentModelNode { Kind = "outline", Text = "Child" };
        var grandchild = new OfficeDocumentModelNode { Kind = "outline", Text = "Grandchild" };
        root.Children = new[] { child };
        child.Children = new[] { grandchild };
        var deepModel = new OfficeDocumentModel { Format = OfficeDocumentFormat.Opml, Structure = new[] { root } };

        Assert.Throws<InvalidDataException>(() => OpmlDocument.FromOfficeDocumentModel(
            deepModel, null, new OpmlConversionOptions { MaxStructureDepth = 2 }));
        Assert.Throws<InvalidDataException>(() => OpmlDocument.FromOfficeDocumentModel(
            deepModel, null, new OpmlConversionOptions { MaxStructureNodes = 2 }));
    }

    [Fact]
    public void SharedForwardConversionBoundsNativeEdits() {
        OpmlDocument deep = OpmlDocument.Create();
        OpmlOutline outline = deep.AddOutline("Root");
        for (int depth = 0; depth < 8; depth++) outline = outline.AddChild("Child");
        Assert.Throws<InvalidDataException>(() => deep.ToOfficeDocumentModel(
            null, new OpmlConversionOptions { MaxStructureDepth = 4 }));

        OpmlDocument wide = OpmlDocument.Create();
        wide.AddOutline("One");
        wide.AddOutline("Two");
        Assert.Throws<InvalidDataException>(() => wide.ToOfficeDocumentModel(
            null, new OpmlConversionOptions { MaxStructureNodes = 1 }));
    }

    [Fact]
    public void SharedConversionDiagnosesUnsupportedMetadata() {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Opml,
            Structure = new[] { new OfficeDocumentModelNode { Kind = "outline", Text = "Root" } },
            Metadata = new[] {
                new OfficeDocumentModelMetadataEntry { Category = "opml", Name = "version", Value = "2.0" },
                new OfficeDocumentModelMetadataEntry { Category = "portable", Name = "subject", Value = "Example" }
            }
        };

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        Assert.Single(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML109" &&
            diagnostic.Message.IndexOf("portable/subject", StringComparison.Ordinal) >= 0);
    }

    [Theory]
    [InlineData("rss", "OPML011")]
    [InlineData("link", "OPML012")]
    [InlineData("include", "OPML012")]
    public void SharedConversionReportsInvalidReconstructedOutlineContracts(string type, string code) {
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Opml,
            Structure = new[] {
                new OfficeDocumentModelNode {
                    Kind = "outline",
                    Text = "Invalid",
                    Attributes = new Dictionary<string, string> { ["type"] = type }
                }
            }
        };

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        Assert.True(converted.HasLoss);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == code && diagnostic.Severity == OpmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void SharedReverseConversionPreservesEditedFlatBlockKind() {
        OfficeDocumentModel model = OpmlDocument.Parse(
            "<opml version=\"2.0\"><head/><body><outline text=\"Original\"/></body></opml>")
            .ToOfficeDocumentModel().Value;
        OfficeDocumentModelBlock block = Assert.Single(model.Blocks);
        block.Kind = "paragraph";

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML107");
        Assert.Equal(2, converted.Value.Outlines.Count);
    }

    [Fact]
    public void SharedReverseConversionReconcilesEditedOutlineTextChannels() {
        OfficeDocumentModel attributeEdited = OpmlDocument.Parse(
            "<opml version=\"2.0\"><head/><body><outline text=\"Original\"/></body></opml>")
            .ToOfficeDocumentModel().Value;
        OfficeDocumentModelNode attributeNode = attributeEdited.Structure.Single();
        Dictionary<string, string> editedAttributes = attributeNode.Attributes.ToDictionary(
            pair => pair.Key, pair => pair.Value, StringComparer.Ordinal);
        editedAttributes["text"] = "Edited attribute";
        attributeNode.Attributes = editedAttributes;

        OpmlConversionResult<OpmlDocument> attributeResult = OpmlDocument.FromOfficeDocumentModel(attributeEdited);

        Assert.Equal("Edited attribute", attributeResult.Value.Outlines.Single().Text);
        Assert.Contains(attributeResult.Diagnostics, diagnostic => diagnostic.Code == "OPML110");

        OfficeDocumentModel primaryEdited = OpmlDocument.Parse(
            "<opml version=\"2.0\"><head/><body><outline text=\"Original\"/></body></opml>")
            .ToOfficeDocumentModel().Value;
        primaryEdited.Structure.Single().Text = "Edited primary";

        OpmlConversionResult<OpmlDocument> primaryResult = OpmlDocument.FromOfficeDocumentModel(primaryEdited);

        Assert.Equal("Edited primary", Assert.Single(primaryResult.Value.Outlines).Text);
        Assert.Contains(primaryResult.Diagnostics, diagnostic => diagnostic.Code == "OPML110");

        OfficeDocumentModel synchronizedPrimary = OpmlDocument.Parse(
            "<opml version=\"2.0\"><head/><body><outline text=\"Original\"/></body></opml>")
            .ToOfficeDocumentModel().Value;
        synchronizedPrimary.Structure.Single().Text = "Edited together";
        synchronizedPrimary.Blocks.Single().Text = "Edited together";

        OpmlConversionResult<OpmlDocument> synchronizedResult = OpmlDocument.FromOfficeDocumentModel(synchronizedPrimary);

        Assert.Equal("Edited together", synchronizedResult.Value.Outlines.Single().Text);
        Assert.Contains(synchronizedResult.Diagnostics, diagnostic => diagnostic.Code == "OPML110" &&
            diagnostic.Message.IndexOf("synchronized", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    [Fact]
    public void SharedReverseConversionSuppressesStaleProjectedLinksAfterRecursiveEdits() {
        OfficeDocumentModel model = OpmlDocument.Parse(
            "<opml version=\"2.0\"><head/><body><outline text=\"Original\" type=\"rss\" url=\"https://example.test/direct\" xmlUrl=\"https://example.test/feed.xml\" htmlUrl=\"https://example.test/\"/></body></opml>")
            .ToOfficeDocumentModel().Value;
        OfficeDocumentModelNode node = Assert.Single(model.Structure);
        node.Text = "Edited";
        Dictionary<string, string> attributes = node.Attributes.ToDictionary(
            pair => pair.Key, pair => pair.Value, StringComparer.Ordinal);
        attributes["url"] = "https://edited.test/direct";
        attributes["xmlUrl"] = "https://edited.test/feed.xml";
        attributes["htmlUrl"] = "https://edited.test/";
        node.Attributes = attributes;

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        OpmlOutline outline = Assert.Single(converted.Value.Outlines);
        Assert.Equal("Edited", outline.Text);
        Assert.Equal("https://edited.test/direct", outline.Url);
        Assert.Equal("https://edited.test/feed.xml", outline.XmlUrl);
        Assert.Equal("https://edited.test/", outline.HtmlUrl);
        Assert.DoesNotContain(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML107");
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SharedReverseConversionRetainsIndependentlyEditedProjectedLinks(bool editUri) {
        OfficeDocumentModel model = OpmlDocument.Parse(
            "<opml version=\"2.0\"><head/><body><outline text=\"Original\" url=\"https://example.test/original\"/></body></opml>")
            .ToOfficeDocumentModel().Value;
        OfficeDocumentModelLink link = Assert.Single(model.Links);
        if (editUri) link.Uri = "https://example.test/edited";
        else link.Text = "Edited flat link";

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        Assert.Equal(2, converted.Value.Outlines.Count);
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML107");
    }

    [Fact]
    public void SharedReverseConversionRetainsIndependentlyEditedFlatBlockText() {
        OfficeDocumentModel model = OpmlDocument.Parse(
            "<opml version=\"2.0\"><head/><body><outline text=\"Original\"/></body></opml>")
            .ToOfficeDocumentModel().Value;
        model.Blocks.Single().Text = "Edited flat block";

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        Assert.Equal(new[] { "Original", "Edited flat block" }, converted.Value.Outlines.Select(outline => outline.Text));
        Assert.Contains(converted.Diagnostics, diagnostic => diagnostic.Code == "OPML107");
    }

    [Theory]
    [InlineData("1.0", false)]
    [InlineData("1.0", true)]
    [InlineData("2.0", false)]
    [InlineData("2.0", true)]
    public void SharedReverseConversionDoesNotRestoreDeletedOutlineProjections(string version, bool removeAll) {
        OfficeDocumentModel model = OpmlDocument.Parse(
            $"<opml version=\"{version}\"><head/><body>" +
            "<outline text=\"Deleted\" url=\"https://example.test/deleted\"/>" +
            "<outline text=\"Retained\" url=\"https://example.test/retained\"/>" +
            "</body></opml>").ToOfficeDocumentModel().Value;
        model.Structure = removeAll ? Array.Empty<OfficeDocumentModelNode>() : model.Structure.Skip(1).ToArray();

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);

        if (removeAll) {
            Assert.Empty(converted.Value.Outlines);
        } else {
            OpmlOutline outline = Assert.Single(converted.Value.Outlines);
            Assert.Equal("Retained", outline.Text);
            Assert.Equal("https://example.test/retained", outline.Url);
        }
        Assert.DoesNotContain(converted.Diagnostics, diagnostic =>
            diagnostic.Code == "OPML101" || diagnostic.Code == "OPML107");
    }

    [Fact]
    public void SharedReverseTraversalDoesNotScheduleChildrenBeyondTheNodeBudget() {
        var children = new WideNodeList();
        var root = new OfficeDocumentModelNode { Kind = "outline", Text = "Root", Children = children };
        var model = new OfficeDocumentModel { Format = OfficeDocumentFormat.Opml, Structure = new[] { root } };

        Assert.Throws<InvalidDataException>(() => OpmlDocument.FromOfficeDocumentModel(
            model, null, new OpmlConversionOptions { MaxStructureNodes = 2 }));
        Assert.Equal(1, children.IndexerCalls);
    }

    private sealed class WideNodeList : IReadOnlyList<OfficeDocumentModelNode> {
        public int Count => int.MaxValue;
        public int IndexerCalls { get; private set; }

        public OfficeDocumentModelNode this[int index] {
            get {
                IndexerCalls++;
                if (index != 0) throw new InvalidOperationException("Traversal accessed a child beyond the configured node budget.");
                return new OfficeDocumentModelNode { Kind = "outline", Text = "First child" };
            }
        }

        public IEnumerator<OfficeDocumentModelNode> GetEnumerator() => throw new NotSupportedException();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
