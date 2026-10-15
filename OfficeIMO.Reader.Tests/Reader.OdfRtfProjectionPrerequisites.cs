using OfficeIMO;
using OfficeIMO.OpenDocument;
using OfficeIMO.Reader;
using OfficeIMO.Reader.OpenDocument;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Rtf;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOdfRtfProjectionPrerequisiteTests {
    [Fact]
    public void LeanPageReconstructionRetainsAnExplicitBreakWithoutLeadingText() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddRtfHandler(new ReaderRtfOptions { IncludePageLocations = true }).Build();
        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.ASCII.GetBytes(@"{\rtf1\page tail}"), "pages.rtf",
            new ReaderOptions { Projection = ReaderProjection.TextAndMetadata });
        Assert.Equal(2, result.Pages.Count); Assert.Empty(result.Pages[0].Blocks);
        Assert.Equal("tail", Assert.Single(result.Pages[1].Blocks).Text);
    }
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void LeanPageReconstructionRetainsCompleteNativeParagraphsIndependentlyOfOutputChunks(bool chunkByBlock) {
        string first = new string('A', 300);
        byte[] source = Encoding.ASCII.GetBytes(@"{\rtf1 " + first + @"\page tail\par next paragraph}");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddRtfHandler(new ReaderRtfOptions {
            ChunkByBlock = chunkByBlock, IncludePageLocations = true
        }).Build();
        OfficeDocumentReadResult result = reader.ReadDocument(source, "pages.rtf", new ReaderOptions {
            Projection = ReaderProjection.TextAndMetadata, MaxChars = 256
        });
        Assert.Equal(2, result.Pages.Count);
        Assert.Equal(first, string.Concat(result.Pages[0].Blocks.Select(block => block.Text)));
        Assert.Contains(result.Pages[1].Blocks, block => block.Text.Contains("tail"));
        Assert.Contains(result.Pages[1].Blocks, block => block.Text.Contains("next paragraph"));
        Assert.Null(result.Markdown); Assert.Empty(result.Assets); Assert.Empty(result.Forms);
    }

    [Fact]
    public void LoadedOpenDocumentProjectionReusesNativeOwnerAndKeepsDeferredPayloadsAvailable() {
        byte[] picture = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        OdtDocument authored = OdtDocument.Create();
        authored.AddParagraph("original").AddImage(picture, "pixel.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        using var input = new MemoryStream(authored.ToBytes());
        OdtDocument loaded = OdtDocument.Load(input, new OdfLoadOptions { DeferBinaryPayloads = true });
        input.Dispose();
        loaded.AddParagraph("edited native text");
        string payloadPath = Assert.Single(loaded.BinaryPayloadEntries);
        var options = new ReaderOptions { Projection = ReaderProjection.TextAndMetadata, MaxInputBytes = 1,
            ResourceLimits = new ReaderResourceLimits { MaxAssets = 0 } };
        OfficeDocumentReadResult projected = loaded.ReadDocument("logical-only.odt", options);
        Assert.Contains(projected.Chunks, chunk => chunk.Text!.Contains("edited native text"));
        Assert.All(projected.Chunks, chunk => Assert.Equal("logical-only.odt", chunk.Location.Path));
        Assert.Null(projected.Source.SourceHash); Assert.Null(projected.Markdown); Assert.Empty(projected.Assets);
        Assert.Equal(OdfPackageEntryPayloadState.Deferred, loaded.GetPackageEntryPayloadState(payloadPath));
        using OfficeEmbeddedPayloadInspection inspection = loaded.InspectEmbeddedPayloads();
        using Stream selected = inspection.OpenRead(0);
        using var bytes = new MemoryStream(); selected.CopyTo(bytes); Assert.Equal(picture, bytes.ToArray());
        Assert.Throws<ReaderResourceLimitException>(() => loaded.ReadDocument("logical-only.odt",
            new ReaderOptions { ResourceLimits = new ReaderResourceLimits { MaxAssets = 0 } }));
        Assert.NotEmpty(loaded.ReadDocument("logical-only.odt").Assets);
        Assert.Equal(ReaderProjection.TextAndMetadata, options.Projection);
    }

    [Fact]
    public void LoadedRtfProjectionPreservesOriginalDiagnosticsAndSourcePayloadAfterSemanticEdits() {
        RtfReadResult loaded = RtfDocument.ParseResult(@"{\rtf1 visible{\pict\pngblip 414243}{\object\objemb{\*\objdata 444546}}}",
            new RtfReadOptions { DeferBinaryPayloads = true });
        Assert.NotEmpty(loaded.Diagnostics);
        loaded.Document.ReplaceText("visible", "edited native text");
        OfficeDocumentReadResult projected = loaded.ReadDocument("logical-only.rtf", new ReaderOptions {
            Projection = ReaderProjection.TextAndMetadata, MaxInputBytes = 1, ResourceLimits = new ReaderResourceLimits { MaxAssets = 0 }
        });
        Assert.Contains(projected.Chunks, chunk => chunk.Text!.Contains("edited native text"));
        Assert.Null(projected.Source.SourceHash); Assert.Null(projected.Markdown); Assert.Empty(projected.Assets);
        foreach (var diagnostic in loaded.Diagnostics) Assert.Contains(projected.Diagnostics, entry => entry.Code == diagnostic.Code);
        using OfficeEmbeddedPayloadInspection inspection = loaded.InspectEmbeddedPayloads();
        Assert.Equal(2, inspection.Payloads.Count);
        using Stream selected = inspection.OpenRead(1);
        using var bytes = new MemoryStream(); selected.CopyTo(bytes); Assert.Equal(Encoding.ASCII.GetBytes("DEF"), bytes.ToArray());
        Assert.Throws<ReaderResourceLimitException>(() => loaded.ReadDocument("logical-only.rtf",
            new ReaderOptions { ResourceLimits = new ReaderResourceLimits { MaxAssets = 0 } }));
        Assert.NotEmpty(loaded.ReadDocument("logical-only.rtf").Assets);
    }

    [Theory]
    [InlineData("odf")]
    [InlineData("rtf")]
    public void LoadedProjectionEnforcesOutputScopeAndCancellationThenRestoresOrdinaryProjection(string format) {
        OdtDocument odf = OdtDocument.Create(); odf.AddParagraph("visible text");
        RtfReadResult rtf = RtfDocument.ParseResult(@"{\rtf1 visible text}");
        OfficeDocumentReadResult Project(ReaderOptions? options = null, CancellationToken token = default) => format == "odf"
            ? odf.ReadDocument("logical.odt", options, cancellationToken: token)
            : rtf.ReadDocument("logical.rtf", options, cancellationToken: token);
        Assert.Throws<ReaderResourceLimitException>(() => Project(new ReaderOptions {
            Projection = ReaderProjection.TextAndMetadata, ResourceLimits = new ReaderResourceLimits { MaxBlocks = 0 }
        }));
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => Project(token: canceled.Token));
        Assert.NotNull(Project().Markdown);
        OfficeDocumentReadResult empty = format == "odf" ? OdtDocument.Create().ReadDocument()
            : RtfDocument.ParseResult(@"{\rtf1}").ReadDocument();
        Assert.Equal(format == "odf" ? "document.odf" : "document.rtf", empty.Source.Path);
    }

    [Theory]
    [InlineData("odf")]
    [InlineData("rtf")]
    public void TextAndMetadataSkipsBinaryAssetsAndMarkdownBeforeAssetBudgetAdmission(string format) {
        OfficeDocumentReader reader;
        byte[] bytes;
        string name;
        if (format == "odf") {
            OdtDocument document = OdtDocument.Create(); document.Metadata.Title = "Source title";
            document.AddParagraph("visible text").AddImage(Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="),
                "pixel.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));
            bytes = document.ToBytes(); name = "source.odt";
            reader = new OfficeDocumentReaderBuilder().AddOpenDocumentHandler().Build();
        } else {
            bytes = Encoding.ASCII.GetBytes(@"{\rtf1{\info{\title Source title}}visible text{\pict\pngblip 414243}}");
            name = "source.rtf"; reader = new OfficeDocumentReaderBuilder().AddRtfHandler().Build();
        }
        OfficeDocumentReadResult lean = reader.ReadDocument(bytes, name, new ReaderOptions {
            Projection = ReaderProjection.TextAndMetadata, ResourceLimits = new ReaderResourceLimits { MaxAssets = 0 }
        });
        Assert.Null(lean.Markdown); Assert.All(lean.Chunks, chunk => Assert.Null(chunk.Markdown));
        Assert.Contains(lean.Chunks, chunk => chunk.Text!.Contains("visible text")); Assert.Equal("Source title", lean.Source.Title);
        Assert.Empty(lean.Assets); Assert.Empty(lean.Visuals); Assert.Empty(lean.Forms); Assert.Empty(lean.OcrCandidates);
        Assert.Contains(lean.Metadata, entry => entry.Value == "not-requested");
        Assert.Contains(reader.GetCapabilityManifest().Handlers, handler => handler.SupportsTextAndMetadata);
        Assert.Throws<ReaderResourceLimitException>(() => reader.ReadDocument(bytes, name,
            new ReaderOptions { ResourceLimits = new ReaderResourceLimits { MaxAssets = 0 } }));
        Assert.NotEmpty(reader.ReadDocument(bytes, name).Assets);
        ReaderChunk[] chunks = reader.Read(bytes, name, new ReaderOptions { Projection = ReaderProjection.TextAndMetadata }).ToArray();
        Assert.All(chunks, chunk => Assert.Null(chunk.Markdown));
    }

    [Theory]
    [InlineData("ods")]
    [InlineData("odp")]
    [InlineData("odg")]
    [InlineData("fodg")]
    public void OpenDocumentProjectionKeepsLocatedTextAcrossRegisteredNativeKinds(string format) {
        OdfDocument document;
        if (format == "ods") { OdsDocument sheets = OdsDocument.Create(); sheets.AddSheet("Data").Cell(0, 0).SetString("visible text"); document = sheets; }
        else if (format == "odp") { OdpPresentation slides = OdpPresentation.Create(); slides.AddSlide("Data").AddTextBox(OdfRect.FromCentimeters(1, 1, 3, 3), "visible text"); document = slides; }
        else { OdgDocument drawing = OdgDocument.Create(); drawing.AddPage().Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 3, 3), "visible text"); document = drawing; }
        byte[] bytes;
        if (format == "fodg") { using var source = new MemoryStream(); document.ToFlatXml().Save(source); bytes = source.ToArray(); }
        else bytes = document.ToBytes();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddOpenDocumentHandler(new ReaderOpenDocumentOptions { HeadersInFirstRow = false }).Build();
        OfficeDocumentReadResult result = reader.ReadDocument(bytes, "source." + format, new ReaderOptions { Projection = ReaderProjection.TextAndMetadata });
        Assert.Contains(result.Chunks, chunk => chunk.Text!.Contains("visible text"));
        Assert.All(result.Chunks, chunk => { Assert.Null(chunk.Markdown); Assert.Equal("source." + format, chunk.Location.Path); });
        Assert.Null(result.Markdown); Assert.Empty(result.Assets); Assert.Empty(result.Visuals);
    }

    [Fact]
    public void UnsupportedDeferredProfilesAreTerminalAndFullDefaultsRemainAvailable() {
        OfficeDocumentReader rtf = new OfficeDocumentReaderBuilder().AddRtfHandler().Build();
        byte[] fromHtml = Encoding.ASCII.GetBytes(@"{\rtf1\fromhtml1 text{\pict\pngblip 414243}}");
        Assert.Throws<NotSupportedException>(() => rtf.ReadDocument(fromHtml, "source.rtf", new ReaderOptions { Projection = ReaderProjection.TextAndMetadata }));
        Assert.NotEmpty(rtf.ReadDocument(fromHtml, "source.rtf").Assets);
        const string password = "test-odf-password";
        OdtDocument document = OdtDocument.Create(); document.AddParagraph("secret").AddImage(Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="),
            "pixel.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        byte[] encrypted = document.ToBytes(new OdfSaveOptions { Encryption = new OdfEncryptionOptions { Password = password } });
        OfficeDocumentReader odf = new OfficeDocumentReaderBuilder().AddOpenDocumentHandler().Build();
        Assert.Equal(OdfEncryptionFailureReason.UnsupportedProfile,
            Assert.Throws<OdfEncryptedPackageException>(() => odf.ReadDocument(encrypted, "source.odt", new ReaderOptions {
                Projection = ReaderProjection.TextAndMetadata, OpenPassword = password
            })).Reason);
        Assert.NotEmpty(odf.ReadDocument(encrypted, "source.odt", new ReaderOptions { OpenPassword = password }).Assets);
        OdgDocument drawing = OdgDocument.Create();
        drawing.AddPage().Shapes.AddImage(Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="),
            "pixel.png", OdfRect.FromCentimeters(1, 1, 1, 1));
        using var flat = new MemoryStream(); drawing.ToFlatXml().Save(flat);
        Assert.Throws<NotSupportedException>(() => odf.ReadDocument(flat.ToArray(), "source.fodg", new ReaderOptions { Projection = ReaderProjection.TextAndMetadata }));
        Assert.NotEmpty(odf.ReadDocument(flat.ToArray(), "source.fodg").Assets);
    }
}
