using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Epub;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubWritingContracts {
    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void Create_Write_ReopensCompleteReadingOrderAndResources(EpubVersion version) {
        EpubPublication book = CreateBook(version);
        Assert.Throws<InvalidOperationException>(() => book.AddSpineItem("first"));
        book.MoveSpineItem(1, 0);
        EpubWriteResult result = book.Write();
        EpubDocument reopened = EpubDocument.Load(new MemoryStream(result.Bytes), new EpubReadOptions { IncludeRawHtml = true, IncludeResourceData = true });
        Assert.Equal("A publishing example", reopened.Title);
        Assert.Equal("Author", reopened.Creator);
        Assert.Equal(new[] { "second", "first" }, reopened.Chapters.Select(chapter => chapter.ManifestId));
        Assert.Equal("Hello, world! Zażółć", reopened.Chapters[1].Text);
        Assert.True(reopened.ReadSummary.IsComplete);
        Assert.Equal(2, reopened.TableOfContents.Count);
        Assert.Contains(reopened.Resources, resource => resource.Id == "style" && resource.Data!.Length > 0);
        Assert.False(result.Report.HasLoss);
        Assert.Equal(result.Bytes, book.Write().Bytes);
        AssertStoredLeadingMimetype(result.Bytes);
    }

    [Fact]
    public void Import_NoEdits_ReturnsExactOriginalPackage() {
        byte[] source = AddEntries(EpubIntegrityFixtures.OneChapter("<p>Original</p>"));
        EpubPublication book = EpubPublication.Load(new MemoryStream(source));
        EpubWriteResult result = book.Write();
        Assert.Equal(source, result.Bytes);
        Assert.True(result.Report.UsedOriginalPackage);
        Assert.Empty(result.Report.RegeneratedEntries);
    }

    [Fact]
    public void Edit_PreservesUnknownXmlUnmanifestedBytesAndUnchangedContent() {
        byte[] source = EpubIntegrityFixtures.OneChapter("<p>Original</p>",
            "<ext:annotation ext:flag='keep'>opaque</ext:annotation>", "xmlns:ext='urn:vendor:extension'");
        source = AddEntries(source, ("vendor/private.bin", new byte[] { 0, 1, 255, 7 }));
        EpubPublication book = EpubPublication.Load(new MemoryStream(source));
        book.Title = "Edited title";
        EpubWriteResult result = book.Write();
        Assert.Equal(ReadEntry(source, "EPUB/chapter.xhtml"), ReadEntry(result.Bytes, "EPUB/chapter.xhtml"));
        Assert.Equal(new byte[] { 0, 1, 255, 7 }, ReadEntry(result.Bytes, "vendor/private.bin"));
        XDocument opf = XDocument.Parse(Encoding.UTF8.GetString(ReadEntry(result.Bytes, "EPUB/package.opf")));
        XElement annotation = Assert.Single(opf.Descendants(XName.Get("annotation", "urn:vendor:extension")));
        Assert.Equal("keep", annotation.Attribute(XName.Get("flag", "urn:vendor:extension"))!.Value);
        Assert.Contains("vendor/private.bin", result.Report.PreservedEntries);
        Assert.Contains("EPUB/package.opf", result.Report.RegeneratedEntries);
        Assert.Equal("Edited title", EpubDocument.Load(new MemoryStream(result.Bytes)).Title);
    }

    [Fact]
    public void ContentXml_TargetedEditRetainsForeignAttributes() {
        EpubPublication book = CreateBook();
        XDocument content = book.GetContentXml("first");
        XNamespace html = "http://www.w3.org/1999/xhtml";
        content.Root!.SetAttributeValue(XName.Get("annotation", "urn:custom"), "keep");
        content.Descendants(html + "p").First().Add(new XElement(html + "span", " additional"));
        book.SetContentXml("first", content);
        Assert.Contains("additional", book.Read().Chapters[0].Text);
        Assert.Equal("keep", book.GetContentXml("first").Root!.Attribute(XName.Get("annotation", "urn:custom"))!.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Save_StagesBeforeReplacingAndLeavesCallerStreamOpen(bool asynchronous) {
        EpubPublication book = CreateBook();
        using var destination = new MemoryStream();
        destination.Write(new byte[100_000], 0, 100_000);
        if (asynchronous) await book.SaveAsync(destination); else book.Save(destination);
        Assert.True(destination.CanWrite);
        Assert.Equal(0, destination.Position);
        Assert.Equal("A publishing example", EpubDocument.Load(destination).Title);
        byte[] before = destination.ToArray();
        Assert.Throws<InvalidDataException>(() => book.Save(destination, new EpubWriteOptions { MaxOutputBytes = 80 }));
        Assert.Equal(before, destination.ToArray());
    }

    [Fact]
    public async Task Save_FileValidationAndCancellationPreserveExistingDestination() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-EpubWriter-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "book.epub");
        try {
            byte[] original = { 1, 2, 3, 4 };
            File.WriteAllBytes(path, original);
            EpubPublication book = CreateBook();
            Assert.Throws<InvalidDataException>(() => book.Save(path, new EpubWriteOptions { MaxExpandedBytes = 20 }));
            Assert.Equal(original, File.ReadAllBytes(path));
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => book.SaveAsync(path, cancellationToken: cancellation.Token));
            Assert.Equal(original, File.ReadAllBytes(path));
            await book.SaveAsync(path);
            Assert.Equal(book.Write().Bytes, File.ReadAllBytes(path));
            Assert.Single(Directory.GetFiles(directory));
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void SignedImport_RequiresExplicitRemovalAndReportsOmission() {
        byte[] source = AddEntries(EpubIntegrityFixtures.OneChapter("<p>Original</p>"),
            ("META-INF/signatures.xml", Encoding.UTF8.GetBytes("<signatures xmlns='urn:oasis:names:tc:opendocument:xmlns:container'/>")));
        EpubPublication book = EpubPublication.Load(new MemoryStream(source));
        Assert.Equal(source, book.Write().Bytes);
        book.Title = "Changed";
        Assert.Throws<InvalidOperationException>(() => book.Write());
        EpubWriteResult result = book.Write(new EpubWriteOptions { RemoveInvalidatedSignatures = true });
        using var archive = new ZipArchive(new MemoryStream(result.Bytes));
        Assert.Null(archive.GetEntry("META-INF/signatures.xml"));
        Assert.True(result.HasLoss);
        Assert.Throws<OfficeConversionException>(() => result.RequireNoLoss());
    }

    [Fact]
    public void ObfuscatedFonts_PreserveOriginalBytesAndRejectIdentityOrPayloadChanges() {
        byte[] source = EpubIntegrityFixtures.Package(new[] {
            ("c", "chapter.xhtml", "application/xhtml+xml", ""), ("font", "font.otf", "font/otf", "") },
            "<itemref idref='c'/>", new[] { ("chapter.xhtml", EpubIntegrityFixtures.Xhtml("<p>Font book</p>")), ("font.otf", "obfuscated-font-bytes") });
        source = AddEntries(source, ("META-INF/encryption.xml", Encoding.UTF8.GetBytes(
            "<encryption xmlns='urn:oasis:names:tc:opendocument:xmlns:container' xmlns:e='http://www.w3.org/2001/04/xmlenc#'>" +
            "<e:EncryptedData><e:EncryptionMethod Algorithm='http://www.idpf.org/2008/embedding'/><e:CipherData>" +
            "<e:CipherReference URI='EPUB/font.otf'/></e:CipherData></e:EncryptedData></encryption>")));
        EpubPublication book = EpubPublication.Load(new MemoryStream(source));
        Assert.Throws<NotSupportedException>(() => book.Identifier = "urn:new:key");
        Assert.Throws<NotSupportedException>(() => book.UpdateResource("font", new byte[] { 1 }));
        book.Title = "New title";
        Assert.Equal(ReadEntry(source, "EPUB/font.otf"), ReadEntry(book.Write().Bytes, "EPUB/font.otf"));
    }

    [Fact]
    public void Navigation_PreservesHierarchyPageListAndLandmarks() {
        EpubPublication book = CreateBook();
        book.SetNavigation(new[] { new EpubNavigationEntry("Part", "EPUB/first.xhtml", new[] {
            new EpubNavigationEntry("Second", "EPUB/second.xhtml#heading") }) },
            new[] { new EpubNavigationEntry("1", "EPUB/second.xhtml#heading") },
            new[] { new EpubNavigationEntry("Start", "EPUB/first.xhtml", semanticType: "bodymatter") });
        EpubDocument reopened = book.Read();
        Assert.Equal("Second", Assert.Single(Assert.Single(reopened.TableOfContents).Children).Label);
        Assert.Single(reopened.PageList);
        Assert.Equal("bodymatter", Assert.Single(reopened.Landmarks).SemanticType);
    }

    [Theory]
    [InlineData("../escape.xhtml")]
    [InlineData("META-INF/private.xhtml")]
    [InlineData("EPUB/invalid:name.xhtml")]
    [InlineData("EPUB/invalid?.xhtml")]
    [InlineData("EPUB/trailing.")]
    public void Authoring_RejectsUnsafeOrReservedPaths(string path) {
        EpubPublication book = CreateBook();
        Assert.Throws<ArgumentException>(() => book.AddResource("invalid", path, "text/plain", new byte[] { 1 }));
    }

    [Fact]
    public void Write_RejectsMissingContentTargetAndFallbackCycleBeforeDestinationChanges() {
        EpubPublication book = CreateBook();
        book.AddChapter("broken", "EPUB/broken.xhtml", "Broken", "<p><img src='missing.png' alt='Missing'/></p>");
        Assert.Throws<InvalidDataException>(() => book.Write());
        book = CreateBook();
        book.Manifest.Single(item => item.Id == "first").FallbackId = "second";
        book.Manifest.Single(item => item.Id == "second").FallbackId = "first";
        Assert.Throws<InvalidDataException>(() => book.Write());
    }

    [Fact]
    public void Load_RejectsRetentionLimitAndPreCanceledStreamsWithoutDisposal() {
        byte[] source = CreateBook().Write().Bytes;
        using var stream = new MemoryStream(source);
        Assert.Throws<EpubReadException>(() => EpubPublication.Load(stream, new EpubPublicationLoadOptions { MaxExpandedBytes = 30 }));
        stream.Position = 0;
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => EpubPublication.Load(stream, cancellationToken: cancellation.Token));
        Assert.Equal(0, stream.Position); Assert.True(stream.CanRead);
    }

    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void Navigation_RejectsMissingFragmentsAndKeepsIdentityConsistent(EpubVersion version) {
        EpubPublication book = CreateBook(version);
        book.Identifier = "urn:book:edited-identity";
        book.SetNavigation(new[] { new EpubNavigationEntry("Second", "EPUB/second.xhtml#heading") },
            pageList: new[] { new EpubNavigationEntry("1", "EPUB/second.xhtml#heading") });
        byte[] bytes = book.Write().Bytes;
        if (version == EpubVersion.Epub2) {
            XDocument ncx = XDocument.Parse(Encoding.UTF8.GetString(ReadEntry(bytes, "EPUB/toc.ncx")));
            Assert.Equal(book.Identifier, ncx.Descendants().Single(element => (string?)element.Attribute("name") == "dtb:uid").Attribute("content")!.Value);
            Assert.Single(ncx.Descendants().Attributes("playOrder").Select(attribute => attribute.Value).Distinct());
            Assert.Equal("1", ncx.Descendants().Single(element => (string?)element.Attribute("name") == "dtb:totalPageCount").Attribute("content")!.Value);
        }
        book.SetNavigation(new[] { new EpubNavigationEntry("Broken", "EPUB/second.xhtml#absent") });
        Assert.Throws<InvalidDataException>(() => book.Write());
    }

    [Fact]
    public void AuthoredContent_DeclaresEmbeddedSvgMathAndRemoteResources() {
        EpubPublication book = CreateBook();
        book.AddChapter("illustrated", "EPUB/illustrated.xhtml", "Illustrated",
            "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><rect width='10' height='10'/></svg>" +
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mi>x</mi></math><img src='https://example.org/image.png' alt='Example'/>" );
        using var output = new MemoryStream(book.Write().Bytes);
        EpubPublication loaded = EpubPublication.Load(output);
        string[] properties = loaded.Manifest.Single(item => item.Id == "illustrated").Properties!.Split(' ');
        Assert.Contains("svg", properties); Assert.Contains("mathml", properties); Assert.Contains("remote-resources", properties);
    }

    [Fact]
    public void ScriptAuthoring_IsRejectedThroughContentAndMutableManifestPaths() {
        EpubPublication book = CreateBook();
        book.AddChapter("script-link", "EPUB/script-link.xhtml", "Script", "<p><a href='javascript:alert(1)'>Run</a></p>");
        Assert.Throws<NotSupportedException>(() => book.Write());
        book = CreateBook();
        book.Manifest.Single(item => item.Id == "first").Properties = "scripted";
        Assert.Throws<NotSupportedException>(() => book.Write());
        Assert.Throws<NotSupportedException>(() => book.AddResource("script", "EPUB/script.js", "application/ecmascript", Encoding.UTF8.GetBytes("alert(1)")));
    }

    [Fact]
    public void RepeatableMetadata_PreservesValuesAndRecognizesVocabularyAliases() {
        EpubPublication book = CreateBook();
        book.DeclareVocabularyPrefix("customSchema", "http://schema.org/");
        book.AddMetadataProperty("customSchema:accessibilityFeature", "structuralNavigation");
        book.AddMetadataProperty("schema:accessibilityFeature", "alternativeText");
        book.SetMetadataProperty("schema:accessMode", "textual");
        EpubPublication loaded = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        XNamespace opf = "http://www.idpf.org/2007/opf";
        Assert.Equal(new[] { "structuralNavigation", "alternativeText" }, loaded.GetPackageXml().Descendants(opf + "meta")
            .Where(element => ((string?)element.Attribute("property"))?.EndsWith(":accessibilityFeature", StringComparison.Ordinal) == true).Select(element => element.Value));
        loaded.SetMetadataProperty("schema:accessibilityFeature", "longDescription");
        Assert.Equal(2, loaded.GetPackageXml().Descendants(opf + "meta").Count(element =>
            ((string?)element.Attribute("property"))?.EndsWith(":accessibilityFeature", StringComparison.Ordinal) == true));
    }

    [Theory]
    [InlineData("pkg-spine-order-svg")]
    [InlineData("pkg-spine-duplicate-item-rendering")]
    public void IndependentW3cImports_EditMetadataPreservesEveryOtherPayload(string name) {
        string directory = Path.Combine(AppContext.BaseDirectory, "w3c", name);
        byte[] source = WriteEntries(Directory.GetFiles(directory, "*", SearchOption.AllDirectories).Select(path =>
            (path.Substring(directory.Length + 1).Replace('\\', '/'), File.ReadAllBytes(path))));
        EpubPublication book = EpubPublication.Load(new MemoryStream(source));
        book.Title = "Edited independent publication";
        EpubWriteResult result = book.Write();
        Assert.All(book.EntryPaths.Where(path => path != book.PackagePath), path => Assert.Equal(ReadEntry(source, path), ReadEntry(result.Bytes, path)));
        Assert.Equal(4, book.Read().Chapters.Count);
        Assert.False(result.HasLoss);
        if (name == "pkg-spine-duplicate-item-rendering") Assert.Contains(result.Report.FidelityDiagnostics,
            diagnostic => diagnostic.Code == "EPUB_WRITE_RETAINED_DUPLICATE_SPINE");
    }

    [Fact]
    public void ResourceRemoval_ReportsOmissionAndProtectsDeclaredCovers() {
        EpubPublication book = CreateBook();
        book.AddResource("extra", "EPUB/extra.bin", "application/octet-stream", new byte[] { 1, 2 });
        book = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        book.RemoveResource("extra");
        EpubWriteResult result = book.Write();
        Assert.Contains("EPUB/extra.bin", result.Report.RemovedEntries);
        Assert.True(result.HasLoss);
        Assert.Throws<OfficeConversionException>(() => result.RequireNoLoss());
        book.AddResource("cover", "EPUB/cover.png", "image/png", new byte[] { 1 });
        book.SetCoverImage("cover");
        Assert.Throws<InvalidOperationException>(() => book.RemoveResource("cover"));
    }

    [Fact]
    public void Load_RejectsBrokenPhysicalMimetypeAndAuthoringRespectsRetainedBudgets() {
        byte[] bytes = CreateBook().Write().Bytes;
        bytes[8] = 8;
        Assert.Throws<InvalidDataException>(() => EpubPublication.Load(new MemoryStream(bytes)));
        EpubPublication book = EpubPublication.Create("Bounded", retentionLimits: new EpubPublicationLoadOptions { MaxEntryBytes = 1024 });
        Assert.Throws<InvalidDataException>(() => book.AddResource("large", "EPUB/large.bin", "application/octet-stream", new byte[1025]));
        Assert.DoesNotContain(book.Manifest, item => item.Id == "large");
    }

    [Fact]
    public void AddChapter_NavigationLimitFailureLeavesThePublicationUnchanged() {
        string title = new string('L', 300);
        EpubPublication probe = EpubPublication.Create("Bounded");
        probe.AddChapter("large-title", "EPUB/title.xhtml", title, "<p>Text</p>");
        int maximumEntryBytes = probe.GetResourceBytes("navigation").Length - 1;
        Assert.True(probe.GetResourceBytes("large-title").Length <= maximumEntryBytes);
        EpubPublication book = EpubPublication.Create("Bounded", retentionLimits: new EpubPublicationLoadOptions { MaxEntryBytes = maximumEntryBytes });
        byte[] originalNavigation = book.GetResourceBytes("navigation");
        Assert.Throws<InvalidDataException>(() => book.AddChapter("large-title", "EPUB/title.xhtml", title, "<p>Text</p>"));
        Assert.Equal(originalNavigation, book.GetResourceBytes("navigation"));
        Assert.Empty(book.Spine);
        Assert.DoesNotContain(book.Manifest, item => item.Id == "large-title");
    }

    [Fact]
    public void UnsupportedCiphertext_PassthroughRetainsBytesButEditingFailsClosed() {
        byte[] source = EpubIntegrityFixtures.ReplaceEntry(EpubIntegrityFixtures.OneChapter("<p>Original</p>"), "EPUB/chapter.xhtml", new byte[] { 0, 255, 1, 7 });
        source = AddEntries(source, ("META-INF/encryption.xml", Encoding.UTF8.GetBytes(
            "<encryption xmlns='urn:oasis:names:tc:opendocument:xmlns:container' xmlns:e='http://www.w3.org/2001/04/xmlenc#'>" +
            "<e:EncryptedData><e:EncryptionMethod Algorithm='urn:unsupported:cipher'/><e:CipherData>" +
            "<e:CipherReference URI='EPUB/chapter.xhtml'/></e:CipherData></e:EncryptedData></encryption>")));
        EpubPublication book = EpubPublication.Load(new MemoryStream(source));
        Assert.Equal(source, book.Write().Bytes);
        book.Title = "Edited";
        Assert.Throws<NotSupportedException>(() => book.Write());
    }

    [Theory]
    [InlineData("oversized")]
    [InlineData("invalid-xml")]
    [InlineData("invalid-uri")]
    [InlineData("duplicate")]
    [InlineData("ambiguous-method")]
    [InlineData("wrong-namespace")]
    public async Task EditableLoad_RejectsUnclassifiedProtectionThroughBothLoadPaths(string failure) {
        string declaration = "<e:EncryptedData><e:EncryptionMethod Algorithm='http://www.idpf.org/2008/embedding'/>" +
            "<e:CipherData><e:CipherReference URI='EPUB/font.otf'/></e:CipherData></e:EncryptedData>";
        string encryption = "<encryption xmlns='urn:oasis:names:tc:opendocument:xmlns:container' xmlns:e='http://www.w3.org/2001/04/xmlenc#'>" +
            declaration + "</encryption>";
        if (failure == "oversized") encryption = encryption.Replace("</encryption>", "<!--" + new string('x', 4096) + "--></encryption>");
        if (failure == "invalid-xml") encryption = "<encryption>";
        if (failure == "invalid-uri") encryption = encryption.Replace("EPUB/font.otf", "https://example.org/font.otf");
        if (failure == "duplicate") encryption = encryption.Replace(declaration, declaration + declaration);
        if (failure == "ambiguous-method") encryption = encryption.Replace("<e:CipherData>", "<e:EncryptionMethod Algorithm='urn:other:cipher'/><e:CipherData>");
        if (failure == "wrong-namespace") encryption = encryption.Replace("http://www.w3.org/2001/04/xmlenc#", "urn:foreign");
        byte[] source = AddEntries(CreateBook().Write().Bytes,
            ("EPUB/font.otf", new byte[] { 1, 2, 3 }), ("META-INF/encryption.xml", Encoding.UTF8.GetBytes(encryption)));
        var limits = new EpubPublicationLoadOptions { MaxMetadataBytes = 4096 };
        Assert.Throws<InvalidDataException>(() => EpubPublication.Load(new MemoryStream(source), limits));
        await Assert.ThrowsAsync<InvalidDataException>(() => EpubPublication.LoadAsync(new MemoryStream(source), limits));
    }

    [Fact]
    public async Task Save_FileCommitFailurePreservesDestinationAndRemovesStagingFiles() {
        // Windows sharing locks prevent replacement; Unix permits renaming an open file.
        if (!System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows)) return;
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-EpubWriter-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "book.epub");
        try {
            byte[] original = { 1, 2, 3, 4 };
            File.WriteAllBytes(path, original);
            using (var held = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                Assert.ThrowsAny<IOException>(() => CreateBook().Save(path));
                await Assert.ThrowsAnyAsync<IOException>(() => CreateBook().SaveAsync(path));
            }
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Single(Directory.GetFiles(directory));
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void ZipDirectorySignature_RequiresExplicitRemovalAndReportsOmission() {
        byte[] source = CreateBook().Write().Bytes;
        int end = source.Length - 22; // This fixture has no ZIP comment.
        Assert.Equal(0x06054b50u, BitConverter.ToUInt32(source, end));
        byte[] signature = { 0x50, 0x4b, 0x05, 0x05, 0x03, 0x00, 0x10, 0x20, 0x30 };
        byte[] signed = new byte[source.Length + signature.Length];
        Buffer.BlockCopy(source, 0, signed, 0, end);
        Buffer.BlockCopy(signature, 0, signed, end, signature.Length);
        Buffer.BlockCopy(source, end, signed, end + signature.Length, 22);
        BitConverter.GetBytes(BitConverter.ToUInt32(source, end + 12) + (uint)signature.Length).CopyTo(signed, end + signature.Length + 12);
        EpubPublication book = EpubPublication.Load(new MemoryStream(signed));
        Assert.Equal(signed, book.Write().Bytes);
        book.Title = "Edited";
        Assert.Throws<InvalidOperationException>(() => book.Write());
        EpubWriteResult result = book.Write(new EpubWriteOptions { RemoveInvalidatedSignatures = true });
        Assert.False(OfficeProvenanceZip.HasCentralDirectorySignature(result.Bytes, 100));
        Assert.True(result.HasLoss);
        Assert.Contains(result.Report.FidelityDiagnostics, item => item.Code == "EPUB_WRITE_ZIP_SIGNATURE_REMOVED");
        Assert.Throws<OfficeConversionException>(() => result.RequireNoLoss());
    }

    [Fact]
    public void ContentEdits_PreserveRemoteDeclarationsForSvgAndExternalStylesheetDependencies() {
        EpubPublication book = CreateBook();
        book.AddResource("illustration", "EPUB/illustration.svg", "image/svg+xml", Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><style>@font-face { font-family: remote; src: url(https://example.org/font.otf); }</style><text x='1' y='5'>Original</text></svg>"), "remote-resources");
        book.AddStylesheet("remote-style", "EPUB/remote.css", "@font-face { font-family: remote; src: url(https://example.org/font.otf); }");
        book.AddChapter("remote-css", "EPUB/remote-css.xhtml", "Remote font", "<p>Original</p>", new[] { "remote-style" });
        book.Manifest.Single(item => item.Id == "remote-css").Properties = "remote-resources";
        book = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        foreach (string id in new[] { "illustration", "remote-css" }) {
            XDocument content = book.GetContentXml(id);
            content.Descendants().First(element => element.Value == "Original").Value = "Edited";
            book.SetContentXml(id, content);
        }
        EpubPublication edited = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.All(new[] { "illustration", "remote-css" }, id =>
            Assert.Contains("remote-resources", edited.Manifest.Single(item => item.Id == id).Properties!.Split(' ')));
    }

    [Fact]
    public void NcxAppend_AllocatesIdsAgainstTheEntireRetainedDocument() {
        EpubPublication book = CreateBook(EpubVersion.Epub2);
        XDocument ncx = XDocument.Parse(Encoding.UTF8.GetString(book.GetResourceBytes("navigation")));
        XNamespace ns = "http://www.daisy.org/z3986/2005/ncx/";
        XElement[] points = ncx.Descendants(ns + "navPoint").ToArray();
        points[0].SetAttributeValue("id", "nav-2"); points[1].SetAttributeValue("id", "nav-3");
        ncx.Root!.Element(ns + "docTitle")!.SetAttributeValue("id", "nav-1");
        book.UpdateResource("navigation", Encoding.UTF8.GetBytes(ncx.ToString()));
        book = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        book.AddChapter("third", "EPUB/third.xhtml", "Third", "<p>Third</p>");
        ncx = XDocument.Parse(Encoding.UTF8.GetString(ReadEntry(book.Write().Bytes, "EPUB/toc.ncx")));
        string[] ids = ncx.Descendants().Attributes("id").Select(attribute => attribute.Value).ToArray();
        Assert.Equal(ids.Length, ids.Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(new[] { "nav-2", "nav-3" }, ncx.Descendants(ns + "navPoint").Take(2).Attributes("id").Select(attribute => attribute.Value));
    }

    [Fact]
    public void NavigationReplacement_PreservesNcxHeadersGuideExtensionsAndHtmlListAttributes() {
        EpubPublication book = CreateBook(EpubVersion.Epub2);
        XNamespace ncxNs = "http://www.daisy.org/z3986/2005/ncx/";
        XNamespace custom = "urn:retained";
        XDocument ncx = XDocument.Parse(Encoding.UTF8.GetString(book.GetResourceBytes("navigation")));
        XElement map = ncx.Root!.Element(ncxNs + "navMap")!;
        map.AddFirst(new XElement(ncxNs + "navInfo", new XAttribute("id", "nav-1"), new XElement(ncxNs + "text", "Retained information")),
            new XElement(ncxNs + "navLabel", new XElement(ncxNs + "text", "Retained label")), new XElement(custom + "extension", "Retained extension"));
        ncx.Root.Add(new XElement(ncxNs + "pageList", new XAttribute("id", "page-1"), new XAttribute("class", "retained-pages"),
            new XElement(ncxNs + "navInfo", new XElement(ncxNs + "text", "Page information")),
            new XElement(ncxNs + "navLabel", new XElement(ncxNs + "text", "Printed pages")), new XElement(custom + "extension", "Page extension")));
        book.UpdateResource("navigation", Encoding.UTF8.GetBytes(ncx.ToString()));
        byte[] source = book.Write().Bytes;
        XDocument package = XDocument.Parse(Encoding.UTF8.GetString(ReadEntry(source, book.PackagePath)));
        XNamespace opf = "http://www.idpf.org/2007/opf";
        package.Root!.Add(new XElement(opf + "guide", new XAttribute(custom + "flag", "retained"), new XElement(custom + "extension", "Guide extension")));
        source = AddEntries(EpubIntegrityFixtures.ReplaceEntry(source, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString())));
        book = EpubPublication.Load(new MemoryStream(source));
        book.SetNavigation(new[] { new EpubNavigationEntry("Replacement", "EPUB/first.xhtml") },
            new[] { new EpubNavigationEntry("2", "EPUB/second.xhtml#heading") },
            new[] { new EpubNavigationEntry("Start", "EPUB/first.xhtml", semanticType: "text") });
        byte[] output = book.Write().Bytes;
        ncx = XDocument.Parse(Encoding.UTF8.GetString(ReadEntry(output, "EPUB/toc.ncx")));
        Assert.Equal("Retained information", ncx.Root!.Element(ncxNs + "navMap")!.Element(ncxNs + "navInfo")!.Value);
        Assert.Equal("Retained label", ncx.Root.Element(ncxNs + "navMap")!.Element(ncxNs + "navLabel")!.Value);
        Assert.Equal("Retained extension", ncx.Root.Element(ncxNs + "navMap")!.Element(custom + "extension")!.Value);
        XElement pages = ncx.Root.Element(ncxNs + "pageList")!;
        Assert.Equal("retained-pages", pages.Attribute("class")!.Value);
        Assert.Equal("Printed pages", pages.Element(ncxNs + "navLabel")!.Value);
        Assert.Equal("Page information", pages.Element(ncxNs + "navInfo")!.Value);
        Assert.Equal("Page extension", pages.Element(custom + "extension")!.Value);
        string[] ids = ncx.Descendants().Attributes("id").Select(attribute => attribute.Value).ToArray();
        Assert.Equal(ids.Length, ids.Distinct(StringComparer.Ordinal).Count());
        XElement guide = XDocument.Parse(Encoding.UTF8.GetString(ReadEntry(output, book.PackagePath))).Root!.Element(opf + "guide")!;
        Assert.Equal("retained", guide.Attribute(custom + "flag")!.Value);
        Assert.Equal("Guide extension", guide.Element(custom + "extension")!.Value);
        Assert.Single(book.Read().TableOfContents); Assert.Single(book.Read().PageList);

        book = CreateBook();
        XDocument html = book.GetContentXml("navigation");
        XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        html.Descendants(xhtml + "ol").First().SetAttributeValue("class", "custom-contents");
        book.SetContentXml("navigation", html);
        book.SetNavigation(new[] { new EpubNavigationEntry("Replacement", "EPUB/first.xhtml") });
        Assert.Equal("custom-contents", book.GetContentXml("navigation").Descendants(xhtml + "ol").First().Attribute("class")!.Value);
    }

    internal static EpubPublication CreateBook(EpubVersion version = EpubVersion.Epub3) {
        EpubPublication book = EpubPublication.Create("A publishing example", "en", "urn:uuid:98c1274a-f755-4ef7-8197-3273a46f1c51", version);
        book.Creator = "Author";
        book.AddStylesheet("style", "EPUB/style.css", "body { color: #093c65; } p { line-height: 1.4; }");
        book.AddChapter("first", "EPUB/first.xhtml", "First", "<p>Hel<em>lo</em>, world<strong>!</strong> Zażółć</p>", new[] { "style" });
        book.AddChapter("second", "EPUB/second.xhtml", "Second", "<h1 id='heading'>Second chapter</h1><p>End.</p>", new[] { "style" });
        return book;
    }
    private static void AssertStoredLeadingMimetype(byte[] data) {
        Assert.Equal(0x04034b50u, BitConverter.ToUInt32(data, 0));
        Assert.Equal(0, BitConverter.ToUInt16(data, 8));
        Assert.Equal(0, BitConverter.ToUInt16(data, 28));
        Assert.Equal("mimetype", Encoding.UTF8.GetString(data, 30, BitConverter.ToUInt16(data, 26)));
    }
    private static byte[] ReadEntry(byte[] package, string path) {
        using var archive = new ZipArchive(new MemoryStream(package));
        using Stream source = archive.GetEntry(path)!.Open();
        using var output = new MemoryStream(); source.CopyTo(output); return output.ToArray();
    }
    private static byte[] AddEntries(byte[] package, params (string Path, byte[] Data)[] added) {
        using var source = new ZipArchive(new MemoryStream(package));
        return WriteEntries(source.Entries.Select(entry => (entry.FullName, ReadEntry(package, entry.FullName))).Concat(added));
    }
    private static byte[] WriteEntries(IEnumerable<(string Path, byte[] Data)> entries) => OfficeProvenanceZipWriter.Write(
        entries.OrderBy(entry => entry.Path == "mimetype" ? 0 : 1).ThenBy(entry => entry.Path, StringComparer.Ordinal)
            .Select(entry => new OfficeProvenanceZipWriteEntry(entry.Path, entry.Data.LongLength, entry.Path != "mimetype",
                new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero), 0, 0, Array.Empty<byte>(), Array.Empty<byte>(), Array.Empty<byte>(),
                () => new MemoryStream(entry.Data, false))).ToArray(), 64L * 1024 * 1024);
}
