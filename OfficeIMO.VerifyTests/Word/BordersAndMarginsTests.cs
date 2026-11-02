using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System.Threading.Tasks;
using VerifyXunit;
using Xunit;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.VerifyTests.Word;

/// <summary>
/// Verifies border and margin related functionality.
/// </summary>
public class BordersAndMarginsTests : VerifyTestBase {

    private static async Task DoTest(WordprocessingDocument document) {
        var result = await ToVerifyResult(document);
        await Verifier.Verify(result, GetSettings());
    }

    [Fact]
    public async Task BasicPageBorders() {
        using var document = WordDocument.Create();
        document.AddParagraph("Section 0");

        document.Sections[0].Borders.LeftStyle = WordBorderStyle.PalmsColor;
        document.Sections[0].Borders.LeftColor = Color.Aqua;
        document.Sections[0].Borders.LeftSpace = OfficeLength.Points(24);
        document.Sections[0].Borders.LeftSize = OfficeLength.Points(3);

        document.Sections[0].Borders.RightStyle = WordBorderStyle.BabyPacifier;
        document.Sections[0].Borders.RightColor = Color.Red;
        document.Sections[0].Borders.RightSize = OfficeLength.Points(1.5);

        document.Sections[0].Borders.TopStyle = WordBorderStyle.SharksTeeth;
        document.Sections[0].Borders.TopColor = Color.GreenYellow;
        document.Sections[0].Borders.TopSize = OfficeLength.Points(1.25);

        document.Sections[0].Borders.BottomStyle = WordBorderStyle.Thick;
        document.Sections[0].Borders.BottomColor = Color.Blue;
        document.Sections[0].Borders.BottomSize = OfficeLength.Points(1.875);

        _ = document.ToBytes();

        await DoTest(document._wordprocessingDocument);
    }

    [Fact]
    public async Task BasicPageBorders2() {
        using var document = WordDocument.Create();
        document.Background.SetColor(Color.DarkSeaGreen);

        document.AddParagraph("Section 0");

        document.Sections[0].SetBorders(WordBorder.Box);

        document.AddSection();
        document.Sections[1].SetBorders(WordBorder.Shadow);

        _ = document.ToBytes();

        await DoTest(document._wordprocessingDocument);
    }

    [Fact]
    public async Task BasicWordMarginsSizes() {
        using var document = WordDocument.Create();
        document.AddParagraph("Section 0");
        document.Sections[0].SetMargins(WordMargin.Normal);

        document.AddSection();
        document.Sections[1].SetMargins(WordMargin.Narrow);
        document.AddParagraph("Section 1");

        document.AddSection();
        document.Sections[2].SetMargins(WordMargin.Mirrored);
        document.AddParagraph("Section 2");

        document.AddSection();
        document.Sections[3].SetMargins(WordMargin.Moderate);
        document.AddParagraph("Section 3");

        document.AddSection();
        document.Sections[4].SetMargins(WordMargin.Wide);
        document.AddParagraph("Section 4");

        _ = document.ToBytes();

        await DoTest(document._wordprocessingDocument);
    }
}
