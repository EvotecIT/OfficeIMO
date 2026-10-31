using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfPdfInheritedFontTests {
    private const string Marker = "iiiiWWWW";

    [Theory]
    [InlineData("paragraph")]
    [InlineData("based-on")]
    [InlineData("character")]
    [InlineData("table")]
    [InlineData("note")]
    [InlineData("attached-note")]
    [InlineData("run-note")]
    [InlineData("field")]
    public void EffectiveStyleFontsMatchDirectFontsInEachContentOwner(string owner) {
        RtfDocument inherited = CreateDocument(owner, directFont: false);
        byte[] source = inherited.ToBytes();
        using PdfCore.PdfDocument expectedModel = CreateDocument(owner, directFont: true).ToPdf(PortableOptions());
        using PdfPigDocument expected = PdfPigDocument.Open(expectedModel.ToBytes());
        using PdfCore.PdfDocument actualModel = inherited.ToPdf(PortableOptions());
        using PdfPigDocument actual = PdfPigDocument.Open(actualModel.ToBytes());

        Assert.Equal(1, actual.NumberOfPages);
        var expectedWord = Assert.Single(expected.GetPage(1).GetWords(), word => word.Text == Marker);
        var actualWord = Assert.Single(actual.GetPage(1).GetWords(), word => word.Text == Marker);
        Assert.Equal(Marker, string.Concat(actualWord.Letters.Select(letter => letter.Value)));
        Assert.All(actualWord.Letters, letter => {
            Assert.Contains("Courier", letter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(12D, letter.PointSize, 3);
        });
        for (int index = 0; index < actualWord.Letters.Count; index++) {
            Assert.Equal(expectedWord.Letters[index].StartBaseLine.X, actualWord.Letters[index].StartBaseLine.X, 3);
            Assert.Equal(expectedWord.Letters[index].StartBaseLine.Y, actualWord.Letters[index].StartBaseLine.Y, 3);
        }
        Assert.Equal(7.2D, actualWord.Letters[1].StartBaseLine.X - actualWord.Letters[0].StartBaseLine.X, 3);
        Assert.Equal(source, inherited.ToBytes());
    }

    [Fact]
    public void EmptyStyledParagraphRegistersTheFontUsedByItsLineBox() {
        RtfDocument inherited = CreateFontDocument(out int courier);
        inherited.AddParagraph("A");
        RtfStyle style = inherited.AddStyle(1, "Empty mono");
        style.FontId = courier;
        style.FontSize = 24;
        inherited.AddParagraph().SetStyle(1);
        inherited.AddParagraph("B");
        byte[] source = inherited.ToBytes();
        PdfCore.PdfDocument actualDocument = inherited.ToPdf(PortableOptions());
        PdfCore.RichParagraphBlock paragraph = Assert.IsType<PdfCore.RichParagraphBlock>(actualDocument.Blocks.ElementAt(1));
        PdfCore.PdfTextRun strut = Assert.Single(paragraph.Runs, run => run.IsLineBoxStrut);
        Assert.Equal(PdfCore.PdfStandardFont.Courier, strut.Font);
        Assert.Equal(24D, strut.FontSize!.Value.ToPoints());

        using PdfPigDocument actual = PdfPigDocument.Open(actualDocument.ToBytes());
        Assert.Equal(1, actual.NumberOfPages);
        Assert.Equal("AB", string.Concat(actual.GetPage(1).Letters.Select(letter => letter.Value)));
        var first = Assert.Single(actual.GetPage(1).Letters, letter => letter.Value == "A");
        var actualLast = Assert.Single(actual.GetPage(1).Letters, letter => letter.Value == "B");
        Assert.True(first.StartBaseLine.Y - actualLast.StartBaseLine.Y > 24D);
        Assert.Equal(source, inherited.ToBytes());
    }

    [Fact]
    public void ControlOnlyStyledTabLeaderMatchesDirectFontGlyphsAndGeometry() {
        using PdfCore.PdfDocument expectedModel = CreateTabLeaderDocument(directFont: true).ToPdf(PortableOptions());
        using PdfPigDocument expected = PdfPigDocument.Open(expectedModel.ToBytes());
        using PdfCore.PdfDocument actualModel = CreateTabLeaderDocument(directFont: false).ToPdf(PortableOptions());
        using PdfPigDocument actual = PdfPigDocument.Open(actualModel.ToBytes());
        Assert.Equal(1, actual.NumberOfPages);
        var expectedDots = expected.GetPage(1).Letters.Where(letter => letter.Value == ".").ToArray();
        var actualDots = actual.GetPage(1).Letters.Where(letter => letter.Value == ".").ToArray();
        Assert.True(actualDots.Length >= 3);
        Assert.Equal(expectedDots.Length, actualDots.Length);
        for (int index = 0; index < actualDots.Length; index++) {
            Assert.Contains("Courier", actualDots[index].FontName, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(expectedDots[index].StartBaseLine.X, actualDots[index].StartBaseLine.X, 3);
            Assert.Equal(expectedDots[index].StartBaseLine.Y, actualDots[index].StartBaseLine.Y, 3);
            Assert.Equal(expectedDots[index].EndBaseLine.X, actualDots[index].EndBaseLine.X, 3);
        }
    }

    [Fact]
    public void UnrelatedRunningStoryDoesNotChangeEffectiveBodyFontOrGeometry() {
        RtfDocument plain = CreateDocument("paragraph", directFont: false);
        RtfDocument withHeader = CreateDocument("paragraph", directFont: false);
        RtfRun header = withHeader.AddHeader().AddParagraph().AddText("HEADER");
        header.FontId = withHeader.Fonts.Single(font => font.Name == "Courier").Id;
        header.FontSize = 12;
        plain.AddParagraph("BODYDEFAULT");
        withHeader.AddParagraph("BODYDEFAULT");

        using PdfCore.PdfDocument expectedModel = plain.ToPdf(PortableOptions());
        using PdfPigDocument expected = PdfPigDocument.Open(expectedModel.ToBytes());
        using PdfCore.PdfDocument actualModel = withHeader.ToPdf(PortableOptions());
        using PdfPigDocument actual = PdfPigDocument.Open(actualModel.ToBytes());
        Assert.Equal(1, actual.NumberOfPages);
        Assert.Single(actual.GetPage(1).GetWords(), word => word.Text == "HEADER");
        foreach (string text in new[] { Marker, "BODYDEFAULT" }) {
            var left = Assert.Single(expected.GetPage(1).GetWords(), word => word.Text == text);
            var right = Assert.Single(actual.GetPage(1).GetWords(), word => word.Text == text);
            Assert.All(right.Letters, letter => Assert.Contains(
                text == Marker ? "Courier" : "Helvetica", letter.FontName, StringComparison.OrdinalIgnoreCase));
            for (int index = 0; index < left.Letters.Count; index++) {
                Assert.Equal(left.Letters[index].FontName, right.Letters[index].FontName);
                Assert.Equal(left.Letters[index].StartBaseLine.X, right.Letters[index].StartBaseLine.X, 3);
                Assert.Equal(left.Letters[index].StartBaseLine.Y, right.Letters[index].StartBaseLine.Y, 3);
            }
        }
    }

    private static RtfDocument CreateDocument(string owner, bool directFont) {
        RtfDocument document = CreateFontDocument(out int courier);
        RtfParagraph paragraph;
        switch (owner) {
            case "table":
                paragraph = document.AddTable(1, 1).Rows[0].Cells[0].AddParagraph();
                break;
            case "note":
            case "attached-note":
            case "run-note":
                RtfNote note = owner == "note" ? document.AddNote(RtfNoteKind.Footnote) : new RtfNote(RtfNoteKind.Footnote);
                RtfParagraph body = document.AddParagraph("BODY");
                if (owner == "run-note") body.AddText("1").SetNote(note);
                else body.AddNoteReference(note, "1");
                paragraph = note.AddParagraph();
                break;
            case "field":
                paragraph = document.AddParagraph().AddField("MERGEFIELD sample").Result;
                break;
            default:
                paragraph = document.AddParagraph();
                break;
        }
        RtfRun run = paragraph.AddText(Marker).SetFontSize(12);
        if (directFont) {
            run.FontId = courier;
        } else if (owner == "character") {
            document.AddStyle(3, "Mono character", RtfStyleKind.Character).FontId = courier;
            run.SetStyle(3);
        } else {
            document.AddStyle(1, "Mono paragraph").FontId = courier;
            if (owner == "based-on") {
                document.AddStyle(2, "Inherited mono").BasedOnStyleId = 1;
                paragraph.SetStyle(2);
            } else {
                paragraph.SetStyle(1);
            }
        }
        return document;
    }

    private static RtfDocument CreateTabLeaderDocument(bool directFont) {
        RtfDocument document = CreateFontDocument(out int courier);
        RtfStyle style = document.AddStyle(1, "Control mono");
        style.FontSize = 12;
        if (!directFont) style.FontId = courier;
        RtfParagraph paragraph = document.AddParagraph().SetStyle(1);
        paragraph.AddTabStop(720, RtfTabAlignment.Left, RtfTabLeader.Dots);
        RtfRun tab = paragraph.AddText("\t");
        if (directFont) tab.FontId = courier;
        return document;
    }

    private static RtfDocument CreateFontDocument(out int courier) {
        RtfDocument document = RtfDocument.Create();
        document.Settings.SetDefaultFont(document.AddFont("Helvetica"));
        courier = document.AddFont("Courier");
        return document;
    }

    private static RtfToPdfOptions PortableOptions() => new RtfToPdfOptions {
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
        TextFallbacks = PdfCore.PdfTextFallbackFeatures.None,
        PdfOptions = new PdfCore.PdfCreateOptions {
            DefaultFont = PdfCore.PdfStandardFont.Helvetica,
            DefaultFontSize = OfficeLength.Points(12)
        }
    };
}
