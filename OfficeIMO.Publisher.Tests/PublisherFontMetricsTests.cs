using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherFontMetricsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Scoped_fonts_survive_options_cloning_projection_and_pdf_export(bool imported) {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Carlito-Regular.ttf"));
        var fonts = new OfficeFontFaceCollection().Add("Arial", font);
        var options = new PublisherReadOptions { Fonts = fonts };
        byte[]? source = imported ? File.ReadAllBytes(PublisherNativeTests.Fixture(
            "NativeComparison/TextVerticalPlacement/native-textverticalplacement.pub")) : null;
        PublisherDocument document = imported ? PublisherDocument.Load(source!, options) : PublisherDocument.Create(200, 100, options);
        // Caller changes must not alter this publication or its future projections.
        fonts.AddAlias("Later alias", "Arial");
        options.Fonts = null;
        if (!imported) document.AddTextFrame(document.Pages[0].Id, 10, 10, 180, 60,
            new[] { new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Scoped font gyjp", 12, OfficeColor.Black, fontFamily: "Arial") }) });
        OfficeDrawing drawing = document.Pages[0].CreateDrawing();
        Assert.Equal("Arial", Assert.Single(drawing.Fonts.Faces).FamilyName);
        string[] stories = document.TextStories.Select(story => story.Text).ToArray();
        byte[] pdfBytes = document.ToPdfBytes();
        PdfReadDocument pdf = PdfReadDocument.Open(pdfBytes);
        var spans = pdf.Pages.SelectMany(page => page.GetTextSpans()).Where(span => !string.IsNullOrWhiteSpace(span.Text)).ToArray();
        Assert.NotEmpty(spans);
        PdfFontInventory inventory = PdfDocument.Load(pdfBytes).Resources.Fonts(new PdfFontInspectionOptions { IncludeEmbeddedProgramBytes = true });
        Assert.NotEmpty(inventory.Fonts);
        Assert.All(inventory.Fonts, item => {
            Assert.True(item.IsEmbedded);
            OfficeTrueTypeFont embedded = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(item.EmbeddedProgramBytes!));
            Assert.Contains("Carlito", embedded.DisplayName ?? "");
        });
        Assert.Contains(imported ? "V0-" : "Scoped font gyjp", pdf.Pages[0].ExtractText());
        Assert.Equal(stories, document.TextStories.Select(story => story.Text));
        if (source != null) Assert.Equal(source, document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        // Returned drawing snapshots are independently mutable too.
        drawing.Fonts.AddAlias("Snapshot alias", "Arial");
        Assert.Single(document.Pages[0].CreateDrawing().Fonts.Faces);
    }
}
