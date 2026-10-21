using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfOptionalBreakLinkTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OptionalBreaks_GeneratedPreflightKeepsSourceOffsetsWithAndWithoutFallbacks(bool fallback) {
        var options = new PdfOptions();
        if (fallback) {
            byte[] font = OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A', 'V');
            options.EmbedStandardFont(PdfStandardFont.Helvetica, font, "IdentifierFont");
            options.RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(new[] {
                new PdfEmbeddedFontFallbackCandidate("OtherIdentifierFont",
                    OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A', 'V', 'B'))
            }, new[] { PdfStandardFont.TimesRoman }));
        }
        string display = fallback ? "AV\u200BBAV" : "AV\u200BAV";
        var document = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text(display));
        Assert.Empty(document.AnalyzeTextEncoding());
        Assert.Equal(fallback ? "AVBAV" : "AVAV", PdfReadDocument.Open(document.ToBytes()).ExtractText().Trim());
        var error = Assert.Single(PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("AV\u200B\uFFFF")).AnalyzeTextEncoding());
        Assert.Equal(3, error.Index);
        Assert.Equal("U+FFFF", error.CodePoint);
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, false, true)]
    [InlineData(true, true, true)]
    public void OptionalBreaks_EmbeddedFontsKeepExactTextWithScalingAndTracking(bool cff, bool named, bool spaced) {
        string? path = cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont() : PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] font = System.IO.File.ReadAllBytes(path!);
        var options = new PdfOptions();
        if (named) { options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("IdentifierFont", font)); }
        else { options.EmbedStandardFont(PdfStandardFont.Helvetica, font, "IdentifierFont"); }
        var run = new PdfTextRun("AV\u200BAV", fontFamily: named ? "IdentifierFont" : null);
        if (spaced) { run = run.WithHorizontalTextScaling(75D).WithCharacterSpacing(0.5D); }

        byte[] bytes = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Runs(new[] { run })).ToBytes();

        Assert.Equal("AVAV", PdfReadDocument.Open(bytes).ExtractText().Trim());
        Assert.Equal(1, PdfInspector.Inspect(bytes).PageCount);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OptionalBreakOnlyLink_KeepsVisibleSiblingAndDestination(bool namedDestination) {
        const string uri = "https://example.test/identifier";
        var document = PdfDocument.Create(new PdfOptions {
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers
        }.UseTextFallbacks());
        document.Paragraph(paragraph => {
            if (namedDestination) {
                paragraph.LinkToBookmark("\u200B", "details", contents: "Separator");
                paragraph.LinkToBookmark("MailHost12", "details", contents: "Visible identifier");
            } else {
                paragraph.Link("\u200B", uri, contents: "Separator");
                paragraph.Link("MailHost12", uri, contents: "Visible identifier");
            }
        });
        document.Bookmark("details").Paragraph(paragraph => paragraph.Text("Destination body"));

        byte[] bytes = document.ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.NotEmpty(info.LinkAnnotations);
        Assert.All(info.LinkAnnotations, link => {
            Assert.True(link.Width > 0);
            Assert.True(link.Height > 0);
            Assert.Equal("Visible identifier", link.Contents);
        });
        Assert.Contains("MailHost12", PdfReadDocument.Open(bytes).ExtractText());
        PdfDocument loaded = PdfDocument.Load(bytes);
        Assert.Equal("details", Assert.Single(loaded.Reader.NamedDestinations()).Name);
        if (namedDestination) { Assert.NotEmpty(loaded.Read().GetLinksByDestinationName("details")); }
        else { Assert.Contains(uri, info.LinkUris); }
    }

    [Fact]
    public void TableOfContents_IdentifierWithOptionalBreaks_KeepsFinalPageAndPositiveLinkAreas() {
        const string display = "Mx12.\u200Blong-\u200Blabel.\u200Bexample.\u200Btest";
        byte[] bytes = PdfDocument.Create(new PdfOptions { TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers })
            .TableOfContents(new PdfTableOfContentsOptions {
                Entries = new[] { new PdfTableOfContentsEntry(display, "details") },
                PageNumberFormatter = page => "page-" + page
            })
            .PageBreak().Bookmark("details").H1("Destination body")
            .ToBytes();

        PdfDocument loaded = PdfDocument.Load(bytes);
        var pages = loaded.Reader.TextByPage();
        Assert.Contains("page-2", pages[0]);
        Assert.Contains("Destination body", pages[1]);
        Assert.Equal(2, Assert.Single(loaded.Reader.NamedDestinations()).PageNumber);
        Assert.NotEmpty(loaded.Read().GetLinksByDestinationName("details"));
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.All(info.LinkAnnotations, link => {
            Assert.True(link.Width > 0);
            Assert.True(link.Height > 0);
        });
        Assert.Contains("Mx12.long-label.example.test", PdfReadDocument.Open(bytes).ExtractText());
    }
}
