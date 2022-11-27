using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using VerifyXunit;
using Xunit;

using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.VerifyTests.Word;

public class BordersAndMarginsTests : VerifyTestBase {

    private static async Task DoTest(WordprocessingDocument document) {
        var result = ToVerifyResult(document);
        await Verifier.Verify(result, GetSettings());
    }

    [Fact]
    public async Task BasicPageBorders() {
        using var document = WordDocument.Create();
        document.AddParagraph("Section 0");

        document.Sections[0].Borders.LeftStyle = BorderValues.PalmsColor;
        document.Sections[0].Borders.LeftColor = Color.Aqua;
        document.Sections[0].Borders.LeftSpace = 24;
        document.Sections[0].Borders.LeftSize = 24;

        document.Sections[0].Borders.RightStyle = BorderValues.BabyPacifier;
        document.Sections[0].Borders.RightColor = Color.Red;
        document.Sections[0].Borders.RightSize = 12;

        document.Sections[0].Borders.TopStyle = BorderValues.SharksTeeth;
        document.Sections[0].Borders.TopColor = Color.GreenYellow;
        document.Sections[0].Borders.TopSize = 10;

        document.Sections[0].Borders.BottomStyle = BorderValues.Thick;
        document.Sections[0].Borders.BottomColor = Color.Blue;
        document.Sections[0].Borders.BottomSize = 15;

        document.Save();

        await DoTest(document._wordprocessingDocument);
    }
}
