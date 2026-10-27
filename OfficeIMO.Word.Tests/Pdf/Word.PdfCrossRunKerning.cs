using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using UglyToad.PdfPig.Content;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfCrossRunKerningTests {
    [Theory]
    [InlineData("body", false)]
    [InlineData("body", true)]
    [InlineData("cell", false)]
    [InlineData("cell", true)]
    [InlineData("header", false)]
    [InlineData("header", true)]
    [InlineData("footer", false)]
    [InlineData("footer", true)]
    public void CompatibleRunBoundariesRetainTheSingleRunGlyphPositions(string route, bool differentColors) {
        Letter[] reference = Render(route, split: false, differentColors: false);
        Letter[] split = Render(route, split: true, differentColors);
        Assert.Equal("AVA", string.Concat(split.Select(letter => letter.Value)));
        for (int index = 0; index < reference.Length; index++) {
            Assert.Equal(reference[index].StartBaseLine.X, split[index].StartBaseLine.X, 3);
            Assert.Equal(reference[index].StartBaseLine.Y, split[index].StartBaseLine.Y, 3);
            if (differentColors) Assert.Equal(index % 2 == 0 ? (0.8D, 0D, 0D) : (0D, 0D, 0.8D), split[index].Color.ToRGBValues());
        }
    }

    private static Letter[] Render(string route, bool split, bool differentColors) {
        using var document = WordDocument.Create();
        document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults =
            new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                new RunFonts { Ascii = ManagedTextShapingTestAssets.FamilyName, HighAnsi = ManagedTextShapingTestAssets.FamilyName },
                new FontSize { Val = "40" }, new Kern { Val = 2 })));
        WordParagraph paragraph = route switch {
            "cell" => document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0],
            "header" => document.HeaderDefaultOrCreate.AddParagraph(),
            "footer" => document.FooterDefaultOrCreate.AddParagraph(),
            _ => document.AddParagraph()
        };
        paragraph._paragraph.RemoveAllChildren<Run>();
        int index = 0;
        foreach (string text in split ? new[] { "A", "V", "A" } : new[] { "AVA" }) {
            var properties = new RunProperties(
                new RunFonts { Ascii = ManagedTextShapingTestAssets.FamilyName, HighAnsi = ManagedTextShapingTestAssets.FamilyName },
                new FontSize { Val = "40" }, new Kern { Val = 2 });
            if (differentColors) properties.Color = new Color { Val = index % 2 == 0 ? "CC0000" : "0000CC" };
            paragraph._paragraph.Append(new Run(properties, new Text(text)));
            index++;
        }
        if (route is "header" or "footer") document.AddParagraph("Body").FontFamily = "Courier";
        var options = new PdfCreateOptions { DefaultFontSize = 20D };
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithPairPositioning('A', 'V')));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, PdfOptions = options, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        return pdf.GetPage(1).Letters.Where(letter => letter.Value is "A" or "V").ToArray();
    }
}
