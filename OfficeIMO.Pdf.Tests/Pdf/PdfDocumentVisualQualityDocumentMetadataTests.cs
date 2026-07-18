using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void ContentStreamFormatting_CanonicalizesNegativeZeroInImageMatrices() {
        byte[] bytes = PdfDocument.Create()
            .Image(CreateMinimalRgbPng(), width: 36, height: 36)
            .ToBytes();

        string stream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));

        Assert.Contains("36 0 0 36", stream, StringComparison.Ordinal);
        Assert.DoesNotContain(" -0 ", stream, StringComparison.Ordinal);
    }

    [Fact]
    public void ContentStreams_CanBeFlateCompressedAndRemainReadable() {
        byte[] uncompressed = CreateCompressionProbe(compressContentStreams: false);
        byte[] compressed = CreateCompressionProbe(compressContentStreams: true);
        string rawCompressed = Encoding.ASCII.GetString(compressed);
        string text = PdfReadDocument.Open(compressed).ExtractText();
        PdfOptions options = new PdfOptions {
            CompressContentStreams = true
        };
        PdfOptions clone = options.Clone();

        Assert.True(compressed.Length < uncompressed.Length, $"Expected compressed PDF to be smaller. Uncompressed: {uncompressed.Length}, compressed: {compressed.Length}.");
        Assert.Contains("/Filter /FlateDecode", rawCompressed, StringComparison.Ordinal);
        Assert.Contains("CompressionProbe", text, StringComparison.Ordinal);
        Assert.Contains("repeated body", text, StringComparison.Ordinal);
        Assert.Equal(1, PdfInspector.Inspect(compressed).PageCount);
        Assert.True(clone.CompressContentStreams);
    }

    [Fact]
    public void StandardFontToUnicodeMaps_CanBeEmittedForGeneratedPdfText() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                IncludeStandardFontToUnicodeMaps = true
            })
            .Paragraph(p => p.Text("Cafe é and Euro €"))
            .ToBytes();
        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Open(bytes).ExtractText();
        PdfOptions clone = new PdfOptions {
            IncludeStandardFontToUnicodeMaps = true
        }.Clone();

        Assert.Contains("/ToUnicode", raw, StringComparison.Ordinal);
        Assert.Contains("/CMapName /OfficeIMO-WinAnsi-UCS", raw, StringComparison.Ordinal);
        Assert.Contains("<80> <20AC>", raw, StringComparison.Ordinal);
        Assert.Contains("<E9> <00E9>", raw, StringComparison.Ordinal);
        Assert.Contains("Cafe é and Euro", text, StringComparison.Ordinal);
        Assert.Contains("€", text, StringComparison.Ordinal);
        Assert.True(clone.IncludeStandardFontToUnicodeMaps);
    }

    [Fact]
    public void XmpMetadata_CanBeEmittedAndSynchronizedWithInfoDictionary() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                IncludeXmpMetadata = true
            })
            .Meta(
                title: "R&D <PDF>",
                author: "OfficeIMO Team",
                subject: "Compliance & metadata",
                keywords: "pdf/a, ua; xmp")
            .Paragraph(p => p.Text("XMP metadata body."))
            .ToBytes();
        string raw = Encoding.UTF8.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(bytes);
        PdfOptions clone = new PdfOptions {
            IncludeXmpMetadata = true
        }.Clone();

        Assert.Contains("/Metadata", raw, StringComparison.Ordinal);
        Assert.Contains("/Type /Metadata /Subtype /XML", raw, StringComparison.Ordinal);
        Assert.Contains("<?xpacket begin=", raw, StringComparison.Ordinal);
        Assert.Contains("<dc:title><rdf:Alt><rdf:li xml:lang=\"x-default\">R&amp;D &lt;PDF&gt;</rdf:li></rdf:Alt></dc:title>", raw, StringComparison.Ordinal);
        Assert.Contains("<dc:creator><rdf:Seq><rdf:li>OfficeIMO Team</rdf:li></rdf:Seq></dc:creator>", raw, StringComparison.Ordinal);
        Assert.Contains("<dc:description><rdf:Alt><rdf:li xml:lang=\"x-default\">Compliance &amp; metadata</rdf:li></rdf:Alt></dc:description>", raw, StringComparison.Ordinal);
        Assert.Contains("<pdf:Keywords>pdf/a, ua; xmp</pdf:Keywords>", raw, StringComparison.Ordinal);
        Assert.Contains("<rdf:li>pdf/a</rdf:li>", raw, StringComparison.Ordinal);
        Assert.Contains("<rdf:li>ua</rdf:li>", raw, StringComparison.Ordinal);
        Assert.Contains("<rdf:li>xmp</rdf:li>", raw, StringComparison.Ordinal);
        Assert.Equal("R&D <PDF>", info.Metadata.Title);
        Assert.Equal("OfficeIMO Team", info.Metadata.Author);
        Assert.Equal("Compliance & metadata", info.Metadata.Subject);
        Assert.Equal("pdf/a, ua; xmp", info.Metadata.Keywords);
        Assert.True(info.HasXmpMetadata);
        Assert.True(preflight.Probe.HasXmpMetadata);
        Assert.True(preflight.CanRewrite);
        Assert.True(clone.IncludeXmpMetadata);
    }

    [Fact]
    public void OutputIntent_CanEmbedIccProfileAndRemainRewriteSafe() {
        byte[] profile = CreateMinimalIccProfile();
        var outputIntent = new PdfOutputIntent(profile, "OfficeIMO RGB") {
            OutputCondition = "OfficeIMO test RGB",
            RegistryName = "https://officeimo.dev/pdf/output-intents",
            Info = "Dependency-free test profile"
        };
        var options = new PdfOptions {
            OutputIntent = outputIntent
        };
        profile[36] = 0;

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("Output intent body."))
            .ToBytes();
        string raw = Encoding.ASCII.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(bytes);
        PdfOptions clone = options.Clone();

        Assert.Contains("/OutputIntents [", raw, StringComparison.Ordinal);
        Assert.Contains("/Type /OutputIntent /S /GTS_PDFA1", raw, StringComparison.Ordinal);
        Assert.Contains("/DestOutputProfile", raw, StringComparison.Ordinal);
        Assert.Contains("/N 3", raw, StringComparison.Ordinal);
        Assert.Contains("<4F6666696365494D4F20524742>", raw, StringComparison.Ordinal);
        Assert.True(info.HasOutputIntents);
        Assert.True(preflight.Probe.HasOutputIntents);
        Assert.True(preflight.CanRewrite);
        Assert.Equal(3, clone.OutputIntent!.ColorComponents);
        Assert.Equal((byte)'a', clone.OutputIntent.IccProfile[36]);
        Assert.NotNull(typeof(PdfOutputIntent).GetConstructor(new[] { typeof(byte[]) }));
        Assert.NotNull(typeof(PdfOutputIntent).GetConstructor(new[] { typeof(byte[]), typeof(string) }));
    }

    [Fact]
    public void OutputIntent_CanUseBuiltInSrgbProfileAndRemainRewriteSafe() {
        var options = new PdfOptions().SetSrgbOutputIntent();
        byte[] mutableProfileCopy = PdfIccProfiles.SrgbIec6196621;
        mutableProfileCopy[36] = 0;

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("Built-in sRGB output intent body."))
            .ToBytes();
        string raw = Encoding.ASCII.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(bytes);
        PdfOutputIntent outputIntent = options.OutputIntent!;

        Assert.Contains("/OutputIntents [", raw, StringComparison.Ordinal);
        Assert.Contains("/Type /OutputIntent /S /GTS_PDFA1", raw, StringComparison.Ordinal);
        Assert.Contains("/DestOutputProfile", raw, StringComparison.Ordinal);
        Assert.Contains("/N 3", raw, StringComparison.Ordinal);
        Assert.Contains("/Length 3052", raw, StringComparison.Ordinal);
        Assert.Contains("<735247422049454336313936362D322E31>", raw, StringComparison.Ordinal);
        Assert.True(info.HasOutputIntents);
        Assert.True(preflight.Probe.HasOutputIntents);
        Assert.True(preflight.CanRewrite);
        Assert.Equal(PdfOutputIntentPolicy.SrgbIec6196621, outputIntent.Policy);
        Assert.Equal(PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier, outputIntent.OutputConditionIdentifier);
        Assert.Equal(3, outputIntent.ColorComponents);
        Assert.Equal((byte)'a', outputIntent.IccProfile[36]);
        Assert.Equal((byte)'a', PdfIccProfiles.SrgbIec6196621[36]);
    }

    [Fact]
    public void OutputIntent_ValidatesIccProfileAndSnapshotsState() {
        byte[] grayProfile = CreateMinimalIccProfile("GRAY");
        byte[] cmykProfile = CreateMinimalIccProfile("CMYK");
        byte[] badSignature = CreateMinimalIccProfile();
        badSignature[36] = (byte)'x';
        byte[] badColorSpace = CreateMinimalIccProfile("LAB ");
        byte[] badDeclaredSize = CreateMinimalIccProfile();
        badDeclaredSize[3] = 131;

        var gray = new PdfOutputIntent(grayProfile, "Gray profile");
        var cmyk = new PdfOutputIntent(cmykProfile, "CMYK profile");
        var options = new PdfOptions().SetSrgbOutputIntent();
        grayProfile[36] = 0;

        Assert.Equal(1, gray.ColorComponents);
        Assert.Equal(4, cmyk.ColorComponents);
        Assert.Equal((byte)'a', options.OutputIntent!.IccProfile[36]);
        Assert.Equal(PdfOutputIntentPolicy.SrgbIec6196621, options.OutputIntent.Policy);
        Assert.Throws<ArgumentException>(() => new PdfOutputIntent(Array.Empty<byte>()));
        Assert.Throws<ArgumentException>(() => new PdfOutputIntent(badSignature));
        Assert.Throws<ArgumentException>(() => new PdfOutputIntent(badColorSpace));
        Assert.Throws<ArgumentException>(() => new PdfOutputIntent(badDeclaredSize));
        Assert.Throws<ArgumentException>(() => new PdfOutputIntent(CreateMinimalIccProfile(), ""));
        Assert.Throws<ArgumentException>(() => new PdfOutputIntent(CreateMinimalIccProfile()) { Info = "" });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfOutputIntent(CreateMinimalIccProfile(), policy: (PdfOutputIntentPolicy)99));
        Assert.NotNull(typeof(PdfOptions).GetMethod("SetOutputIntent", new[] { typeof(byte[]), typeof(string) }));
        Assert.NotNull(typeof(PdfDocument).GetMethod("OutputIntent", new[] { typeof(byte[]), typeof(string) }));
    }

    [Fact]
    public void DocumentLanguage_CanBeEmittedAndInspected() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                Language = "en-US"
            })
            .Paragraph(p => p.Text("Language body."))
            .ToBytes();
        string raw = Encoding.ASCII.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(bytes);
        PdfOptions clone = new PdfOptions {
            Language = "pl-PL"
        }.Clone();

        Assert.Contains("/Lang <656E2D5553>", raw, StringComparison.Ordinal);
        Assert.Equal("en-US", info.CatalogLanguage);
        Assert.Equal("en-US", preflight.DocumentInfo!.CatalogLanguage);
        Assert.Equal("pl-PL", clone.Language);
        Assert.Throws<ArgumentException>(() => new PdfOptions { Language = "" });
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().Language("bad\u0001lang"));
    }

    [Fact]
    public void PageLabels_CanBeEmittedAndInspected() {
        var options = new PdfOptions {
            IncludePageLabels = true,
            PageNumberStyle = PdfPageNumberStyle.UpperRoman,
            PageNumberStart = 3,
            PageLabelPrefix = "A-"
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("Page label proof."))
            .PageBreak()
            .Paragraph(p => p.Text("Second labelled page."))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(bytes);
        PdfPageLabel label = Assert.Single(info.PageLabels);
        PdfOptions clone = options.Clone();

        Assert.Contains("/PageLabels ", raw, StringComparison.Ordinal);
        Assert.Contains("/S /R", raw, StringComparison.Ordinal);
        Assert.Contains("/St 3", raw, StringComparison.Ordinal);
        Assert.Contains("/P <412D>", raw, StringComparison.Ordinal);
        Assert.True(info.HasPageLabels);
        Assert.True(info.HasReadablePageLabels);
        Assert.True(preflight.Probe.HasPageLabels);
        Assert.True(preflight.CanRewrite);
        Assert.Equal(0, label.StartPageIndex);
        Assert.Equal(1, label.StartPageNumber);
        Assert.Equal("R", label.Style);
        Assert.Equal("A-", label.Prefix);
        Assert.Equal(3, label.StartNumber);
        Assert.True(clone.IncludePageLabels);
        Assert.Equal("A-", clone.PageLabelPrefix);

        byte[] extracted = PdfPageExtractor.ExtractPages(bytes, 2);
        PdfPageLabel extractedLabel = Assert.Single(PdfInspector.Inspect(extracted).PageLabels);
        Assert.Equal(0, extractedLabel.StartPageIndex);
        Assert.Equal("A-", extractedLabel.Prefix);
        Assert.Equal(4, extractedLabel.StartNumber);

        Assert.Throws<ArgumentException>(() => new PdfOptions { PageLabelPrefix = "" });
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().PageLabels("bad\u0001prefix"));
    }

    [Fact]
    public void PageLabelRanges_CanBeEmittedClonedInspectedAndReindexed() {
        var options = new PdfOptions()
            .AddPageLabelRange(3, PdfPageNumberStyle.Arabic, 1)
            .AddPageLabelRange(1, PdfPageNumberStyle.LowerRoman, 1, "front-");

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("Cover."))
            .PageBreak()
            .Paragraph(p => p.Text("Contents."))
            .PageBreak()
            .Paragraph(p => p.Text("Chapter one."))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(bytes);
        PdfOptions clone = options.Clone();

        Assert.Contains("/PageLabels ", raw, StringComparison.Ordinal);
        Assert.Contains("/Nums [0 << /S /r /St 1 /P <66726F6E742D> >> 2 << /S /D /St 1 >>]", raw, StringComparison.Ordinal);
        Assert.True(info.HasPageLabels);
        Assert.True(info.HasReadablePageLabels);
        Assert.True(preflight.CanRewrite);
        Assert.Equal(2, info.PageLabels.Count);
        Assert.Equal(new[] { 0, 2 }, info.PageLabels.Select(label => label.StartPageIndex).ToArray());
        Assert.Equal(new[] { "r", "D" }, info.PageLabels.Select(label => label.Style).ToArray());
        Assert.Equal(new[] { "front-", null }, info.PageLabels.Select(label => label.Prefix).ToArray());
        Assert.Equal(new[] { 1, 1 }, info.PageLabels.Select(label => label.StartNumber!.Value).ToArray());
        Assert.True(clone.IncludePageLabels);
        Assert.Equal(2, clone.PageLabelRanges.Count);
        Assert.Equal(new[] { 1, 3 }, clone.PageLabelRanges.Select(range => range.StartPageNumber).ToArray());

        byte[] extracted = PdfPageExtractor.ExtractPages(bytes, 2, 3);
        PdfDocumentInfo extractedInfo = PdfInspector.Inspect(extracted);
        Assert.Equal(2, extractedInfo.PageLabels.Count);
        Assert.Equal(new[] { 0, 1 }, extractedInfo.PageLabels.Select(label => label.StartPageIndex).ToArray());
        Assert.Equal(new[] { "front-", null }, extractedInfo.PageLabels.Select(label => label.Prefix).ToArray());
        Assert.Equal(new[] { 2, 1 }, extractedInfo.PageLabels.Select(label => label.StartNumber!.Value).ToArray());

        Assert.Throws<ArgumentException>(() => options.AddPageLabelRange(1, PdfPageNumberStyle.Arabic));
        Assert.Throws<InvalidOperationException>(() =>
            PdfDocument.Create().PageLabelRange(2, PdfPageNumberStyle.Arabic).Paragraph(p => p.Text("Only one page.")).ToBytes());
    }

    [Fact]
    public void CatalogView_CanBeEmittedClonedInspectedAndPreservedOnExtraction() {
        var options = new PdfOptions()
            .SetCatalogView(PdfCatalogPageMode.FullScreen, PdfCatalogPageLayout.TwoColumnLeft);

        byte[] bytes = PdfDocument.Create(options)
            .CatalogPageMode(PdfCatalogPageMode.UseThumbs)
            .CatalogPageLayout(PdfCatalogPageLayout.SinglePage)
            .Paragraph(p => p.Text("Catalog view proof."))
            .PageBreak()
            .Paragraph(p => p.Text("Second page."))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(bytes);
        PdfOptions clone = options.Clone();

        Assert.Contains("/PageMode /UseThumbs", raw, StringComparison.Ordinal);
        Assert.Contains("/PageLayout /SinglePage", raw, StringComparison.Ordinal);
        Assert.True(preflight.Probe.HasCatalogViewSettings);
        Assert.True(preflight.CanRewrite);
        Assert.Equal("UseThumbs", info.CatalogPageMode);
        Assert.Equal("SinglePage", info.CatalogPageLayout);
        Assert.Equal(PdfCatalogPageMode.FullScreen, clone.CatalogPageMode);
        Assert.Equal(PdfCatalogPageLayout.TwoColumnLeft, clone.CatalogPageLayout);

        byte[] extracted = PdfPageExtractor.ExtractPages(bytes, 1);
        PdfDocumentInfo extractedInfo = PdfInspector.Inspect(extracted);
        Assert.Equal("UseThumbs", extractedInfo.CatalogPageMode);
        Assert.Equal("SinglePage", extractedInfo.CatalogPageLayout);

        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfOptions { CatalogPageMode = (PdfCatalogPageMode)99 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfOptions { CatalogPageLayout = (PdfCatalogPageLayout)99 });
    }

    [Fact]
    public void OpenAction_CanBeEmittedClonedInspectedAndFilteredOnExtraction() {
        var options = new PdfOptions().SetOpenAction(pageNumber: 2, destinationTop: 700D);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("First page."))
            .PageBreak()
            .Paragraph(p => p.Text("Open here."))
            .PageBreak()
            .Paragraph(p => p.Text("Third page."))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(bytes);
        PdfDocumentOpenAction openAction = Assert.IsType<PdfDocumentOpenAction>(info.OpenAction);
        PdfOpenActionOptions cloneOpenAction = Assert.IsType<PdfOpenActionOptions>(options.Clone().OpenAction);

        Assert.Contains("/OpenAction [", raw, StringComparison.Ordinal);
        Assert.Contains("/XYZ 0 700 0", raw, StringComparison.Ordinal);
        Assert.True(info.HasOpenActions);
        Assert.True(info.HasReadableOpenAction);
        Assert.True(preflight.Probe.HasOpenActions);
        Assert.True(preflight.CanRewrite);
        Assert.Equal("Destination", openAction.ActionType);
        Assert.Equal(2, openAction.PageNumber);
        Assert.Equal(700D, openAction.DestinationTop);
        Assert.Equal(PdfOpenActionDestinationMode.Xyz, openAction.DestinationMode);
        Assert.Equal(2, cloneOpenAction.PageNumber);
        Assert.Equal(700D, cloneOpenAction.DestinationTop);
        Assert.Equal(PdfOpenActionDestinationMode.Xyz, cloneOpenAction.DestinationMode);

        byte[] extractedTarget = PdfPageExtractor.ExtractPages(bytes, 2);
        PdfDocumentOpenAction extractedOpenAction = Assert.IsType<PdfDocumentOpenAction>(PdfInspector.Inspect(extractedTarget).OpenAction);
        Assert.Equal(1, extractedOpenAction.PageNumber);
        Assert.Equal(700D, extractedOpenAction.DestinationTop);
        Assert.Equal(PdfOpenActionDestinationMode.Xyz, extractedOpenAction.DestinationMode);

        byte[] extractedOther = PdfPageExtractor.ExtractPages(bytes, 1);
        PdfDocumentInfo otherInfo = PdfInspector.Inspect(extractedOther);
        Assert.False(otherInfo.HasOpenActions);
        Assert.False(otherInfo.HasReadableOpenAction);

        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfOpenActionOptions(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfOpenActionOptions(1, double.PositiveInfinity));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfOpenActionOptions(1, destinationMode: (PdfOpenActionDestinationMode)99));
        Assert.Throws<InvalidOperationException>(() =>
            PdfDocument.Create().OpenAction(2).Paragraph(p => p.Text("Only one page.")).ToBytes());
    }

    [Fact]
    public void OpenAction_CanEmitFitAndFitHorizontalDestinations() {
        byte[] fitBytes = PdfDocument.Create(new PdfOptions().SetOpenAction(destinationMode: PdfOpenActionDestinationMode.Fit))
            .Paragraph(p => p.Text("Fit page."))
            .ToBytes();
        string fitRaw = Encoding.ASCII.GetString(fitBytes);
        PdfDocumentOpenAction fitOpenAction = Assert.IsType<PdfDocumentOpenAction>(PdfInspector.Inspect(fitBytes).OpenAction);

        Assert.Contains("/OpenAction [", fitRaw, StringComparison.Ordinal);
        Assert.Contains("/Fit]", fitRaw, StringComparison.Ordinal);
        Assert.DoesNotContain("/XYZ", fitRaw, StringComparison.Ordinal);
        Assert.Equal(PdfOpenActionDestinationMode.Fit, fitOpenAction.DestinationMode);
        Assert.Equal(1, fitOpenAction.PageNumber);
        Assert.Null(fitOpenAction.DestinationTop);

        byte[] fitHorizontalBytes = PdfDocument.Create()
            .OpenAction(1, 640D, PdfOpenActionDestinationMode.FitHorizontal)
            .Paragraph(p => p.Text("Fit horizontal page."))
            .ToBytes();
        string fitHorizontalRaw = Encoding.ASCII.GetString(fitHorizontalBytes);
        PdfDocumentOpenAction fitHorizontalOpenAction = Assert.IsType<PdfDocumentOpenAction>(PdfInspector.Inspect(fitHorizontalBytes).OpenAction);

        Assert.Contains("/FitH 640]", fitHorizontalRaw, StringComparison.Ordinal);
        Assert.Equal(PdfOpenActionDestinationMode.FitHorizontal, fitHorizontalOpenAction.DestinationMode);
        Assert.Equal(1, fitHorizontalOpenAction.PageNumber);
        Assert.Equal(640D, fitHorizontalOpenAction.DestinationTop);

        byte[] extracted = PdfPageExtractor.ExtractPages(fitHorizontalBytes, 1);
        PdfDocumentOpenAction extractedOpenAction = Assert.IsType<PdfDocumentOpenAction>(PdfInspector.Inspect(extracted).OpenAction);
        Assert.Equal(PdfOpenActionDestinationMode.FitHorizontal, extractedOpenAction.DestinationMode);
        Assert.Equal(640D, extractedOpenAction.DestinationTop);
    }

    [Fact]
    public void OpenAction_CanEmitExtendedDestinationModes() {
        byte[] fitVerticalBytes = PdfDocument.Create()
            .OpenAction(1, destinationMode: PdfOpenActionDestinationMode.FitVertical, destinationLeft: 36D)
            .Paragraph(p => p.Text("Fit vertical page."))
            .ToBytes();
        string fitVerticalRaw = Encoding.ASCII.GetString(fitVerticalBytes);
        PdfDocumentOpenAction fitVerticalOpenAction = Assert.IsType<PdfDocumentOpenAction>(PdfInspector.Inspect(fitVerticalBytes).OpenAction);
        Assert.Contains("/FitV 36]", fitVerticalRaw, StringComparison.Ordinal);
        Assert.Equal(PdfOpenActionDestinationMode.FitVertical, fitVerticalOpenAction.DestinationMode);
        Assert.Equal(36D, fitVerticalOpenAction.DestinationLeft);
        Assert.Null(fitVerticalOpenAction.DestinationTop);

        byte[] wideFitVerticalBytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 200D,
                PageHeight = 200D
            })
            .OpenAction(1, destinationMode: PdfOpenActionDestinationMode.FitVertical, destinationLeft: 300D)
            .Paragraph(p => p.Text("Fit vertical can use one coordinate."))
            .ToBytes();
        PdfDocumentOpenAction wideFitVerticalOpenAction = Assert.IsType<PdfDocumentOpenAction>(PdfInspector.Inspect(wideFitVerticalBytes).OpenAction);
        Assert.Contains("/FitV 300]", Encoding.ASCII.GetString(wideFitVerticalBytes), StringComparison.Ordinal);
        Assert.Equal(PdfOpenActionDestinationMode.FitVertical, wideFitVerticalOpenAction.DestinationMode);
        Assert.Equal(300D, wideFitVerticalOpenAction.DestinationLeft);

        byte[] fitRectangleBytes = PdfDocument.Create()
            .OpenAction(1, 640D, PdfOpenActionDestinationMode.FitRectangle, destinationLeft: 24D, destinationBottom: 48D, destinationRight: 320D)
            .Paragraph(p => p.Text("Fit rectangle page."))
            .ToBytes();
        string fitRectangleRaw = Encoding.ASCII.GetString(fitRectangleBytes);
        PdfDocumentOpenAction fitRectangleOpenAction = Assert.IsType<PdfDocumentOpenAction>(PdfInspector.Inspect(fitRectangleBytes).OpenAction);
        Assert.Contains("/FitR 24 48 320 640]", fitRectangleRaw, StringComparison.Ordinal);
        Assert.Equal(PdfOpenActionDestinationMode.FitRectangle, fitRectangleOpenAction.DestinationMode);
        Assert.Equal(24D, fitRectangleOpenAction.DestinationLeft);
        Assert.Equal(48D, fitRectangleOpenAction.DestinationBottom);
        Assert.Equal(320D, fitRectangleOpenAction.DestinationRight);
        Assert.Equal(640D, fitRectangleOpenAction.DestinationTop);

        byte[] fitBoundingBoxBytes = PdfDocument.Create(new PdfOptions().SetOpenAction(destinationMode: PdfOpenActionDestinationMode.FitBoundingBox))
            .Paragraph(p => p.Text("Fit bounding box page."))
            .ToBytes();
        string fitBoundingBoxRaw = Encoding.ASCII.GetString(fitBoundingBoxBytes);
        PdfDocumentOpenAction fitBoundingBoxOpenAction = Assert.IsType<PdfDocumentOpenAction>(PdfInspector.Inspect(fitBoundingBoxBytes).OpenAction);
        Assert.Contains("/FitB]", fitBoundingBoxRaw, StringComparison.Ordinal);
        Assert.Equal(PdfOpenActionDestinationMode.FitBoundingBox, fitBoundingBoxOpenAction.DestinationMode);

        byte[] fitBoundingHorizontalBytes = PdfDocument.Create()
            .OpenAction(1, 610D, PdfOpenActionDestinationMode.FitBoundingBoxHorizontal)
            .Paragraph(p => p.Text("Fit bounding box horizontal page."))
            .ToBytes();
        PdfDocumentOpenAction fitBoundingHorizontalOpenAction = Assert.IsType<PdfDocumentOpenAction>(PdfInspector.Inspect(fitBoundingHorizontalBytes).OpenAction);
        Assert.Contains("/FitBH 610]", Encoding.ASCII.GetString(fitBoundingHorizontalBytes), StringComparison.Ordinal);
        Assert.Equal(PdfOpenActionDestinationMode.FitBoundingBoxHorizontal, fitBoundingHorizontalOpenAction.DestinationMode);
        Assert.Equal(610D, fitBoundingHorizontalOpenAction.DestinationTop);

        byte[] fitBoundingVerticalBytes = PdfDocument.Create()
            .OpenAction(1, destinationMode: PdfOpenActionDestinationMode.FitBoundingBoxVertical, destinationLeft: 42D)
            .Paragraph(p => p.Text("Fit bounding box vertical page."))
            .ToBytes();
        PdfDocumentOpenAction fitBoundingVerticalOpenAction = Assert.IsType<PdfDocumentOpenAction>(PdfInspector.Inspect(fitBoundingVerticalBytes).OpenAction);
        Assert.Contains("/FitBV 42]", Encoding.ASCII.GetString(fitBoundingVerticalBytes), StringComparison.Ordinal);
        Assert.Equal(PdfOpenActionDestinationMode.FitBoundingBoxVertical, fitBoundingVerticalOpenAction.DestinationMode);
        Assert.Equal(42D, fitBoundingVerticalOpenAction.DestinationLeft);
    }

    [Fact]
    public void EmbeddedFiles_CanBeEmittedAsNameTreeAndAssociatedFiles() {
        byte[] invoiceXml = Encoding.UTF8.GetBytes("<rsm:CrossIndustryInvoice>42</rsm:CrossIndustryInvoice>");
        byte[] sourceText = Encoding.UTF8.GetBytes("Generated from OfficeIMO");

        byte[] bytes = PdfDocument.Create(new PdfOptions()
                .AddEmbeddedFile("invoice.xml", invoiceXml, "application/xml", PdfAssociatedFileRelationship.Data, "Structured invoice XML"))
            .AttachFile("source.txt", sourceText, "text/plain", PdfAssociatedFileRelationship.Source)
            .Paragraph(p => p.Text("Embedded file proof."))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Names << /EmbeddedFiles", raw, StringComparison.Ordinal);
        Assert.Contains("/AF [", raw, StringComparison.Ordinal);
        Assert.Contains("/Type /Filespec", raw, StringComparison.Ordinal);
        Assert.Contains("/Type /EmbeddedFile", raw, StringComparison.Ordinal);
        Assert.Contains("/AFRelationship /Data", raw, StringComparison.Ordinal);
        Assert.Contains("/AFRelationship /Source", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /application#2Fxml", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /text#2Fplain", raw, StringComparison.Ordinal);
        Assert.Contains("/Params << /Size 55 /CheckSum <83F5425DBE5CB56CCCFFC5F749EDCAD4> >>", raw, StringComparison.Ordinal);
        Assert.Contains("/Params << /Size 24 /CheckSum <C321A4BD26D4D3AF6F40C2FC4EDDF6AE> >>", raw, StringComparison.Ordinal);
        Assert.Contains("CrossIndustryInvoice", raw, StringComparison.Ordinal);

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(bytes);
        Assert.True(info.HasEmbeddedFiles);
        Assert.True(preflight.Probe.HasEmbeddedFiles);
        Assert.True(preflight.CanRewrite);

        byte[] extracted = PdfPageExtractor.ExtractPages(bytes, 1);
        Assert.Contains("/EmbeddedFiles", Encoding.ASCII.GetString(extracted), StringComparison.Ordinal);
    }

    [Fact]
    public void EmbeddedFiles_SnapshotDataAndRejectInvalidInputs() {
        byte[] data = { 1, 2, 3 };
        var file = new PdfEmbeddedFile("note.txt", data, "text/plain", PdfAssociatedFileRelationship.Supplement, "Note");
        data[0] = 9;

        byte[] snapshot = file.Data;
        snapshot[1] = 9;

        Assert.Equal(1, file.Data[0]);
        Assert.Equal(2, file.Data[1]);
        Assert.Equal("Supplement", PdfEmbeddedFileDictionaryBuilder.GetRelationshipName(file.Relationship));

        var options = new PdfOptions().AddEmbeddedFile(file);
        file.FileName = "changed.txt";
        PdfEmbeddedFile stored = Assert.Single(options.EmbeddedFiles);
        stored.FileName = "snapshot.txt";

        Assert.Equal("note.txt", Assert.Single(options.EmbeddedFiles).FileName);
        Assert.Equal("note.txt", Assert.Single(options.Clone().EmbeddedFiles).FileName);

        options.ClearEmbeddedFiles();
        Assert.Empty(options.EmbeddedFiles);

        Assert.Throws<ArgumentNullException>(() => new PdfOptions().AddEmbeddedFile(null!));
        Assert.Throws<ArgumentException>(() => new PdfOptions()
            .AddEmbeddedFile("note.txt", new byte[] { 1 })
            .AddEmbeddedFile("note.txt", new byte[] { 2 }));
        Assert.Throws<ArgumentException>(() => new PdfEmbeddedFile("folder/note.txt", new byte[] { 1 }));
        Assert.Throws<ArgumentException>(() => new PdfEmbeddedFile("note.txt", Array.Empty<byte>()));
        Assert.Throws<ArgumentException>(() => new PdfEmbeddedFile("note.txt", new byte[] { 1 }, "text plain"));
        Assert.Throws<ArgumentException>(() => new PdfEmbeddedFile("note.txt", new byte[] { 1 }) { Description = "" });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfEmbeddedFile("note.txt", new byte[] { 1 }, relationship: (PdfAssociatedFileRelationship)99));
    }

    [Fact]
    public void FileVersion_CanEmitPdf17HeaderAndCloneOptions() {
        var options = new PdfOptions {
            FileVersion = PdfFileVersion.Pdf17
        };
        PdfOptions clone = options.Clone();

        byte[] bytes = PdfDocument.Create()
            .FileVersion(PdfFileVersion.Pdf17)
            .Paragraph(p => p.Text("PDF 1.7 header proof."))
            .ToBytes();

        Assert.Equal(PdfFileVersion.Pdf17, clone.FileVersion);
        Assert.StartsWith("%PDF-1.7", Encoding.ASCII.GetString(bytes));
        Assert.Equal("1.7", PdfInspector.Inspect(bytes).HeaderVersion);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfOptions {
            FileVersion = (PdfFileVersion)99
        });
    }

    [Fact]
    public void ComplianceProfile_ValidatesAndClonesOptions() {
        var options = new PdfOptions {
            FileVersion = PdfFileVersion.Pdf17,
            ComplianceProfile = PdfComplianceProfile.PdfA3U,
            IncludeXmpMetadata = true,
            IncludeStandardFontToUnicodeMaps = true
        };
        PdfOptions clone = options.Clone();

        var invalidException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfOptions {
                ComplianceProfile = (PdfComplianceProfile)999
            });

        Assert.Equal(PdfFileVersion.Pdf17, clone.FileVersion);
        Assert.Equal(PdfComplianceProfile.PdfA3U, clone.ComplianceProfile);
        Assert.True(clone.IncludeXmpMetadata);
        Assert.True(clone.IncludeStandardFontToUnicodeMaps);
        Assert.Contains("PDF compliance profile must be None", invalidException.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(PdfComplianceProfile.PdfA2U, "PDF/A-2u", "Unicode text mapping", "veraPDF")]
    [InlineData(PdfComplianceProfile.PdfA2A, "PDF/A-2a", "tagged PDF structure tree", "alternate text")]
    [InlineData(PdfComplianceProfile.PdfA3U, "PDF/A-3u", "Unicode text mapping", "veraPDF")]
    [InlineData(PdfComplianceProfile.PdfA3A, "PDF/A-3a", "tagged PDF structure tree", "alternate text")]
    public void ComplianceProfile_RejectsUnsupportedFormalProfiles(PdfComplianceProfile profile, string displayName, string requirement, string validator) {
        var exception = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Create()
                .Compliance(profile)
                .Meta(title: "Compliance probe")
                .Paragraph(p => p.Text("Body"))
                .ToBytes());

        Assert.Contains(displayName, exception.Message, StringComparison.Ordinal);
        Assert.Contains("cannot yet generate certified", exception.Message, StringComparison.Ordinal);
        Assert.Contains(requirement, exception.Message, StringComparison.Ordinal);
        Assert.Contains(validator, exception.Message, StringComparison.Ordinal);
        Assert.Contains(nameof(PdfComplianceProfile.None), exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ComplianceProfile_PdfAAccessibilityMessageDoesNotRequirePdfUaSpecificMetadata() {
        var exception = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Create()
                .Compliance(PdfComplianceProfile.PdfA3A)
                .Paragraph(p => p.Text("Body"))
                .ToBytes());

        Assert.Contains("PDF/A-3a", exception.Message, StringComparison.Ordinal);
        Assert.Contains("tagged PDF structure tree", exception.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("PDF/UA identification XMP", exception.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("document title metadata", exception.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("DisplayDocTitle", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void EmbeddedStandardFonts_SnapshotDataAndRejectInvalidInputs() {
        var data = new byte[] { 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        var options = new PdfOptions()
            .EmbedStandardFont(PdfStandardFont.Helvetica, data, "Snapshot Font");
        data[0] = 255;
        PdfEmbeddedFont embeddedFont = options.EmbeddedFonts[PdfStandardFont.Helvetica];
        byte[] readback = embeddedFont.Data;
        readback[1] = 255;
        PdfOptions clone = options.Clone();
        var renderException = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Create(options)
                .Paragraph(p => p.Text("Invalid embedded font"))
                .ToBytes());

        Assert.Equal(0, embeddedFont.Data[0]);
        Assert.Equal(1, embeddedFont.Data[1]);
        Assert.Equal("Snapshot Font", clone.EmbeddedFonts[PdfStandardFont.Helvetica].FontName);
        Assert.True(clone.CompressEmbeddedFonts);
        Assert.Throws<ArgumentException>(() => new PdfOptions().EmbedStandardFont(PdfStandardFont.Helvetica, Array.Empty<byte>()));
        Assert.Contains("TrueType font", renderException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void EmbeddedStandardFonts_CanWriteTrueTypeFontFileResourcesWhenAvailable() {
        string? fontPath = FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressEmbeddedFonts = true
            })
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMOEmbeddedArial")
            .Paragraph(p => p.Text("Embedded font Cafe é and Euro €"))
            .ToBytes();
        byte[] uncompressed = PdfDocument.Create(new PdfOptions {
                CompressEmbeddedFonts = false
            })
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMOEmbeddedArial")
            .Paragraph(p => p.Text("Embedded font Cafe é and Euro €"))
            .ToBytes();
        string raw = Encoding.ASCII.GetString(bytes);
        string rawUncompressed = Encoding.ASCII.GetString(uncompressed);
        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.True(bytes.Length < uncompressed.Length, $"Expected compressed embedded font PDF to be smaller. Compressed: {bytes.Length}, uncompressed: {uncompressed.Length}.");
        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /CIDFontType2", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /OfficeIMOEmbeddedArial", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("/CIDToGIDMap /Identity", raw, StringComparison.Ordinal);
        Assert.Contains("/FontDescriptor", raw, StringComparison.Ordinal);
        Assert.Contains("/FontFile2", raw, StringComparison.Ordinal);
        AssertSubsetLength1(raw, fontData.Length);
        Assert.Contains("/Filter /FlateDecode", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/Filter /FlateDecode", rawUncompressed, StringComparison.Ordinal);
        Assert.Contains("/W [", raw, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", raw, StringComparison.Ordinal);
        Assert.Contains("Embedded font Cafe é and Euro", text, StringComparison.Ordinal);
        Assert.Contains("€", text, StringComparison.Ordinal);
    }

    [Fact]
    public void UseFontFamily_EmbedsNamedTrueTypeFamilyForGeneratedText() {
        string? fontPath = FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressEmbeddedFonts = true,
                CompressContentStreams = false
            })
            .UseFontFamily("OfficeIMO Pretty", fontData)
            .Header(header => header.Text("Pretty header"))
            .Paragraph(paragraph => paragraph
                .Text("Pretty regular ")
                .Bold("pretty bold ")
                .Italic()
                .Text("pretty italic"))
            .Footer(footer => footer.Text("Pretty footer {page}/{pages}"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /CIDFontType2", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /OfficeIMOPretty-Regular", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /OfficeIMOPretty-Bold", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /OfficeIMOPretty-Italic", raw, StringComparison.Ordinal);
        Assert.Contains("/FontFile2", raw, StringComparison.Ordinal);
        AssertSubsetLength1(raw, fontData.Length);
        Assert.Contains("Pretty regular pretty bold pretty italic", text, StringComparison.Ordinal);
        Assert.Contains("Pretty header", text, StringComparison.Ordinal);
        Assert.Contains("Pretty footer 1/1", text, StringComparison.Ordinal);
    }

    private static void AssertSubsetLength1(string raw, int originalFontLength) {
        MatchCollection matches = Regex.Matches(raw, @"/Length1\s+(\d+)");
        Assert.NotEmpty(matches);
        foreach (Match match in matches) {
            int length = int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
            Assert.InRange(length, 1, originalFontLength - 1);
        }
    }


}
