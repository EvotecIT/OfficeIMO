using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("image", 24D)]
    [InlineData("shape", 36D)]
    [InlineData("group", 36D)]
    [InlineData("hidden-shape", 36D)]
    public void SaveAsPdf_ObjectOnlyParagraphDoesNotAddAnotherEmptyLine(string kind, double objectHeight) {
        double baseline = ObjectMarkGap("none", false);
        double actual = ObjectMarkGap(kind, false);
        Assert.Equal(objectHeight, actual - baseline, 3);
        double withBlank = ObjectMarkGap(kind, true);
        Assert.Equal(baseline, withBlank - actual, 3);
    }

    [Fact]
    public void SaveAsPdf_ObjectWithoutFlowHeightStillRetainsItsVisibleParagraphMark() {
        double baseline = ObjectMarkGap("none", false);
        Assert.Equal(baseline * 2D, ObjectMarkGap("anchored-image", false), 3);
    }

    [Theory]
    [InlineData("image", false)]
    [InlineData("image", true)]
    [InlineData("shape", false)]
    [InlineData("shape", true)]
    [InlineData("group", false)]
    [InlineData("group", true)]
    public void SaveAsPdf_ObjectParagraphRetainsDirectAndInheritedSpacingAndCollapsesNeighbors(string kind, bool inherited) {
        double baseline = ObjectMarkGap(kind, false);
        Assert.Equal(40D, ObjectMarkGap(kind, false, inherited, spaced: true) - baseline, 3);
        Assert.Equal(60D, ObjectMarkGap(kind, false, inherited, spaced: true, neighborSpacing: 30D) - baseline, 3);
    }

    private static double ObjectMarkGap(string kind, bool followingBlank, bool inherited = false, bool spaced = false, double neighborSpacing = 0D) {
        using WordDocument document = WordDocument.Create();
        W.Style normal = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<W.Style>().Single(s => s.StyleId?.Value == "Normal");
        normal.StyleRunProperties = new W.StyleRunProperties(
            new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
            new W.FontSize { Val = "24" }, new W.FontSizeComplexScript { Val = "24" });
        normal.StyleParagraphProperties = new W.StyleParagraphProperties(new W.SpacingBetweenLines {
            Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto
        });
        document.AddParagraph("A").LineSpacingAfterPoints = neighborSpacing;
        if (kind != "none") {
            WordParagraph p = document.AddParagraph();
            if (spaced) {
                if (inherited) {
                    document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart!.Styles!.Append(new W.Style {
                        StyleId = "ObjectSpacing", Type = W.StyleValues.Paragraph, CustomStyle = true,
                        StyleName = new W.StyleName { Val = "Object Spacing" }, BasedOn = new W.BasedOn { Val = "Normal" },
                        StyleParagraphProperties = new W.StyleParagraphProperties(new W.SpacingBetweenLines { Before = "400", After = "400" })
                    });
                    p.SetStyleId("ObjectSpacing");
                } else {
                    p.LineSpacingBeforePoints = 20;
                    p.LineSpacingAfterPoints = 20;
                }
            }
            if (kind == "group") {
                p.AddShapeGroup(new[] {
                    new WordShapeGroupItem(WordShapeType.Rectangle, 0, 0, 24, 36) { FillColorHex = "00A060" },
                    new WordShapeGroupItem(WordShapeType.Rectangle, 24, 0, 24, 36) { FillColorHex = "00A060" }
                });
            } else if (kind is "shape" or "hidden-shape") {
                WordShape shape = p.AddShape(48, 36, "#00A060");
                if (kind == "hidden-shape") shape.Hidden = true;
            } else {
                using var image = new MemoryStream(OfficeRasterImageEncoder.Encode(
                    new OfficeRasterImage(64, 32, OfficeColor.Green), OfficeImageExportFormat.Png));
                WordImage inserted = p.InsertImage(image, "mark.png", 64, 32,
                    kind == "anchored-image" ? WordImageTextWrapping.InFrontOfText : WordImageTextWrapping.InLineWithText);
                if (kind == "anchored-image") {
                    inserted.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
                    inserted.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
                    inserted.HorizontalPositionOffset = 140 * 12700;
                    inserted.VerticalPositionOffset = 160 * 12700;
                }
            }
            if (followingBlank) document.AddParagraph();
        }
        document.AddParagraph("B").LineSpacingBeforePoints = neighborSpacing;
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        var letters = pdf.GetPage(1).Letters;
        return Assert.Single(letters, l => l.Value == "A").StartBaseLine.Y -
            Assert.Single(letters, l => l.Value == "B").StartBaseLine.Y;
    }
}
