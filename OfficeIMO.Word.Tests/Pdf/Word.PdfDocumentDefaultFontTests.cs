using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_DocumentDefaultFontKeepsItsRegisteredFamilyAlongsideFallbacks(bool withFallbacks, bool explicitBody) {
        const string family = "OfficeIMO document default";
        using WordDocument document = WordDocument.Create();
        var defaults = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults ??= new DocDefaults();
        defaults.RunPropertiesDefault ??= new RunPropertiesDefault();
        defaults.RunPropertiesDefault.RunPropertiesBaseStyle ??= new RunPropertiesBaseStyle();
        defaults.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts = new RunFonts { Ascii = family, HighAnsi = family };
        document.AddParagraph("Default");

        var options = new PdfOptions();
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family, ManagedTextShapingTestAssets.CreateFont('D', 'e', 'f', 'a', 'u', 'l', 't')));
        if (withFallbacks) {
            options.RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(new[] {
                new PdfEmbeddedFontFallbackCandidate("Fallback", ManagedTextShapingTestAssets.CreateFont('D', 'e', 'f', 'a', 'u', 'l', 't'))
            }));
        }
        if (explicitBody) options.DefaultFont = PdfStandardFont.Courier;

        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options
        }));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToList();
        Assert.Equal("Default", string.Concat(letters.Select(letter => letter.Value)));
        // An explicit PDF body font still wins over the document default.
        if (explicitBody) Assert.All(letters, letter => Assert.Contains("Courier", letter.FontName));
        else Assert.All(letters, letter => Assert.DoesNotContain("Helvetica", letter.FontName));
    }
}
