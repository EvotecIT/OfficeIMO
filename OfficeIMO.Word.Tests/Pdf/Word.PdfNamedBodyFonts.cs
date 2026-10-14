using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, "Arial", 0)]
    [InlineData(true, "Arial", 0)]
    [InlineData(false, "Times New Roman", 0)]
    [InlineData(true, "Times New Roman", 0)]
    [InlineData(false, "Named Source, Times New Roman", 0)]
    [InlineData(true, "Named Source, Times New Roman", 0)]
    [InlineData(false, "Arial", 1)]
    [InlineData(true, "Arial", 1)]
    [InlineData(false, "Times New Roman", 1)]
    [InlineData(true, "Times New Roman", 1)]
    [InlineData(false, "Named Source, Times New Roman", 1)]
    [InlineData(true, "Named Source, Times New Roman", 1)]
    [InlineData(false, "Arial", 2)]
    [InlineData(true, "Arial", 2)]
    [InlineData(false, "Times New Roman", 2)]
    [InlineData(true, "Times New Roman", 2)]
    [InlineData(false, "Named Source, Times New Roman", 2)]
    [InlineData(true, "Named Source, Times New Roman", 2)]
    public void SaveAsPdf_InheritedNamedBodyUsesRegisteredResource(bool profile, string sourceFamily, int bodyOverride) {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = sourceFamily;
        document.Settings.FontSize = 18;
        document.AddParagraph("BODY");
        var options = new WordToPdfOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        if (profile) options.UseRenderingProfile(new OfficeRenderingProfile("shaping-only"));
        else options.PdfOptions = new PdfOptions();
        string registeredFamily = sourceFamily.StartsWith("Named Source", StringComparison.Ordinal) ? "Named Source" : sourceFamily;
        options.PdfOptions!.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(registeredFamily,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray())));
        if (bodyOverride == 1) options.PdfOptions.DefaultFont = PdfStandardFont.Courier;
        if (bodyOverride == 2) options.FontFamily = "Courier New";

        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(options));
        Assert.Equal("BODY", pdf.GetPage(1).Text);
        Assert.All(pdf.GetPage(1).Letters, letter => {
            Assert.Contains(bodyOverride == 0 ? registeredFamily.Replace(" ", string.Empty) + "-Regular" : "Courier", letter.FontName, StringComparison.Ordinal);
            Assert.Equal(18D, letter.PointSize, 3);
        });
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public void SaveAsPdf_FontBearingProfileAndClonePreserveBodyIntent(bool overlay, bool matchedBody, bool explicitBody) {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = matchedBody ? "Arial" : "Times New Roman";
        document.Settings.FontSize = 18;
        document.AddParagraph("BODY");
        var profile = new OfficeRenderingProfile("named-profile", new OfficeFontFaceCollection()
            .Add("Arial", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray())));
        var options = new WordToPdfOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() }.UseRenderingProfile(profile);
        if (overlay) options.UseRenderingProfile(new OfficeRenderingProfile("fontless-overlay"), OfficeRenderingProfileApplyMode.Overlay);
        if (explicitBody) options.PdfOptions!.DefaultFont = PdfStandardFont.Courier;
        options.PdfOptions!.HeaderFontFamily = "Arial";
        options.PdfOptions.ShowHeader = true;
        options.PdfOptions.HeaderFormat = "HEADER";
        using var original = PdfPigDocument.Open(document.ToPdfBytes(options));
        using var cloned = PdfPigDocument.Open(document.ToPdfBytes(options.Clone()));
        string expected = explicitBody ? "Courier" : matchedBody ? "Arial-Regular" : "Times";
        var body = Assert.Single(cloned.GetPage(1).GetWords(), word => word.Text == "BODY").Letters;
        Assert.All(body, letter => {
            Assert.Contains(expected, letter.FontName, StringComparison.Ordinal);
            Assert.Equal(18D, letter.PointSize, 3);
        });
        var header = Assert.Single(cloned.GetPage(1).GetWords(), word => word.Text == "HEADER").Letters;
        Assert.All(header, letter => Assert.Contains("Arial-Regular", letter.FontName, StringComparison.Ordinal));
        Assert.Equal(original.GetPage(1).Letters.Select(letter => (letter.Value, letter.FontName, letter.PointSize, letter.StartBaseLine)),
            cloned.GetPage(1).Letters.Select(letter => (letter.Value, letter.FontName, letter.PointSize, letter.StartBaseLine)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_ClearingStandardSlotsRestoresBodyDiscoveryWithNamedResources(bool profile) {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = "Times New Roman";
        document.Settings.FontSize = 18;
        document.AddParagraph("BODY");
        WordToPdfOptions options = CreateNamedPresentationOptions(profile);
        options.PdfOptions!.RegisterFontFamily(PdfStandardFont.Helvetica,
            new PdfEmbeddedFontFamily("Body Slot", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray())));
        using var configured = PdfPigDocument.Open(document.ToPdfBytes(options));
        Assert.All(configured.GetPage(1).Letters, letter => Assert.Contains("BodySlot-Regular", letter.FontName, StringComparison.Ordinal));
        options.PdfOptions.ClearEmbeddedStandardFonts();
        using var discovered = PdfPigDocument.Open(document.ToPdfBytes(options));
        Assert.All(discovered.GetPage(1).Letters, letter => {
            Assert.Contains("Times", letter.FontName, StringComparison.Ordinal);
            Assert.Equal(18D, letter.PointSize, 3);
        });
        options.PdfOptions.ClearNamedFontFamilies();
        using var cleared = PdfPigDocument.Open(document.ToPdfBytes(options));
        Assert.Equal(discovered.GetPage(1).Letters.Select(letter => (letter.Value, letter.FontName, letter.PointSize, letter.StartBaseLine)),
            cleared.GetPage(1).Letters.Select(letter => (letter.Value, letter.FontName, letter.PointSize, letter.StartBaseLine)));
    }
}
