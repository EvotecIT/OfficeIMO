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

    [Fact]
    public void LimitsAndDtdPolicyRejectHostileInputs() {
        Assert.Throws<InvalidDataException>(() => OpmlDocument.Parse(
            "<opml version=\"2.0\"><head/><body><outline text=\"a\"/><outline text=\"b\"/></body></opml>",
            new OpmlReadOptions { MaxOutlines = 1 }));
        Assert.ThrowsAny<Exception>(() => OpmlDocument.Parse(
            "<!DOCTYPE opml [<!ENTITY x \"boom\">]><opml version=\"2.0\"><head/><body><outline text=\"&x;\"/></body></opml>"));
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => OpmlDocument.Parse("<opml version=\"2.0\"><head/><body/></opml>", cancellationToken: canceled.Token));
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
        OpmlOutline root = source.AddOutline("Root"); root.SetAttribute(XName.Get("flag", "urn:test"), "yes"); root.AddChild("Child");
        var model = source.ToOfficeDocumentModel().Value;
        Assert.Equal(OfficeDocumentFormat.Opml, model.Format);
        Assert.Equal("Child", model.Structure.Single().Children.Single().Text);
        Assert.Equal(2, model.Structure.SelectMany(node => new[] { node }.Concat(node.Children)).Select(node => node.Id).Distinct().Count());

        OpmlConversionResult<OpmlDocument> converted = OpmlDocument.FromOfficeDocumentModel(model);
        Assert.False(converted.HasLoss);
        Assert.Equal(OpmlVersion.Opml10, converted.Value.Version);
        Assert.Equal("yes", converted.Value.Outlines.Single().GetAttribute(XName.Get("flag", "urn:test")));
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
        OpmlDocument document = OpmlDocument.Parse("<opml version=\"2.0\" custom=\"x\">root-text<head>head-text</head><body>body-text<!--native--><outline text=\"A\">outline-text<extra/></outline><body-extension/></body><root-extension/></opml>");
        OpmlConversionResult<OfficeDocumentModel> result = document.ToOfficeDocumentModel();
        Assert.True(result.HasLoss);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML200");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML202");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML203");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML204");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML205");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML206");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OPML207");
    }
}
