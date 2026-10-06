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

    [Theory]
    [InlineData("image", false)]
    [InlineData("image", true)]
    [InlineData("shape", false)]
    [InlineData("shape", true)]
    [InlineData("group", false)]
    [InlineData("group", true)]
    public void SaveAsPdf_ObjectParagraphHonorsDirectAndInheritedMinimumLineHeight(string kind, bool inherited) {
        Assert.Equal(80D, ObjectMarkGap(kind, false, inherited, minimumHeight: 80D) - ObjectMarkGap("none", false), 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_MixedFlowAndAnchoredImagesUseOneParagraphFrame(bool spaced) {
        Assert.Equal(ObjectMarkGap("image", false, spaced: spaced),
            ObjectMarkGap("image", false, spaced: spaced, mixedAnchor: true), 3);
    }

    [Theory]
    [InlineData("image", true)]
    [InlineData("image", false)]
    [InlineData("shape", true)]
    [InlineData("shape", false)]
    [InlineData("group", true)]
    [InlineData("group", false)]
    public void SaveAsPdf_ObjectParagraphCollapsesSpacingWithVisibleBlankNeighbors(string kind, bool blankBefore) {
        double baseline = ObjectMarkGap(kind, false, spaced: true, neighborSpacing: 30D);
        double actual = ObjectMarkGap(kind, !blankBefore, spaced: true, neighborSpacing: 30D, spacedBlank: true, blankBefore: blankBefore);
        Assert.Equal(ObjectMarkGap("none", false) + 20D, actual - baseline, 3);
    }

    [Theory]
    [InlineData("image-short")]
    [InlineData("shape-short")]
    [InlineData("group-short")]
    public void SaveAsPdf_ShortObjectRetainsItsNaturalParagraphMarkMinimum(string kind) {
        double markHeight = ObjectMarkGap("none", false);
        Assert.Equal(markHeight * 2D, ObjectMarkGap(kind, false), 3);
    }

    [Theory]
    [InlineData("shape", true)]
    [InlineData("shape-short", true)]
    [InlineData("shape-inline", false)]
    [InlineData("shape-inline-short", false)]
    public void SaveAsPdf_MinimumLineHeightRetainsShapeLineAlignment(string kind, bool lineTop) {
        double normalTop = 0D, minimumTop = 0D;
        double normalGap = ObjectMarkGap(kind, false, inspect: page => normalTop = ShapeTop(page));
        ObjectMarkGap(kind, false, minimumHeight: 80D, inspect: page => minimumTop = ShapeTop(page));
        double expectedShift = lineTop ? 0D : 80D - (normalGap - ObjectMarkGap("none", false));
        Assert.Equal(expectedShift, normalTop - minimumTop, 3);

        static double ShapeTop(UglyToad.PdfPig.Content.Page page) =>
            Assert.Single(page.Paths, path => path.IsFilled).GetBoundingRectangle()!.Value.Top;
    }

    private static double ObjectMarkGap(string kind, bool followingBlank, bool inherited = false, bool spaced = false,
        double neighborSpacing = 0D, double minimumHeight = 0D, bool mixedAnchor = false, bool spacedBlank = false, bool blankBefore = false,
        Action<UglyToad.PdfPig.Content.Page>? inspect = null) {
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
            void AddBlank() {
                WordParagraph blank = document.AddParagraph();
                if (spacedBlank) { blank.LineSpacingBeforePoints = 20D; blank.LineSpacingAfterPoints = 20D; }
            }
            if (blankBefore) AddBlank();
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
            if (minimumHeight > 0D) {
                var spacing = new W.SpacingBetweenLines { Line = (minimumHeight * 20D).ToString(System.Globalization.CultureInfo.InvariantCulture), LineRule = W.LineSpacingRuleValues.AtLeast };
                if (inherited) {
                    document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart!.Styles!.Append(new W.Style {
                        StyleId = "ObjectMinimum", Type = W.StyleValues.Paragraph, CustomStyle = true,
                        StyleName = new W.StyleName { Val = "Object Minimum" }, BasedOn = new W.BasedOn { Val = "Normal" },
                        StyleParagraphProperties = new W.StyleParagraphProperties(spacing)
                    });
                    p.SetStyleId("ObjectMinimum");
                } else {
                    p._paragraph.ParagraphProperties ??= new W.ParagraphProperties();
                    p._paragraph.ParagraphProperties.SpacingBetweenLines = spacing;
                }
            }
            if (kind.StartsWith("group", StringComparison.Ordinal)) {
                p.AddShapeGroup(new[] {
                    new WordShapeGroupItem(WordShapeType.Rectangle, 0, 0, 24, kind.EndsWith("short", StringComparison.Ordinal) ? 6 : 36) { FillColorHex = "00A060" },
                    new WordShapeGroupItem(WordShapeType.Rectangle, 24, 0, 24, kind.EndsWith("short", StringComparison.Ordinal) ? 6 : 36) { FillColorHex = "00A060" }
                });
            } else if (kind.StartsWith("shape-inline", StringComparison.Ordinal)) {
                WordShape shape = WordShape.AddDrawingShape(p, WordShapeType.Rectangle, 48, kind.EndsWith("short", StringComparison.Ordinal) ? 6 : 36);
                shape.FillColorHex = "00A060";
            } else if (kind.StartsWith("shape", StringComparison.Ordinal) || kind == "hidden-shape") {
                WordShape shape = p.AddShape(48, kind.EndsWith("short", StringComparison.Ordinal) ? 6 : 36, "#00A060");
                if (kind == "hidden-shape") shape.Hidden = true;
            } else {
                using var image = new MemoryStream(OfficeRasterImageEncoder.Encode(
                    new OfficeRasterImage(64, 32, OfficeColor.Green), OfficeImageExportFormat.Png));
                WordImage inserted = p.InsertImage(image, "mark.png", 64, kind == "image-short" ? 8 : 32,
                    kind == "anchored-image" ? WordImageTextWrapping.InFrontOfText : WordImageTextWrapping.InLineWithText);
                if (kind == "anchored-image") {
                    inserted.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
                    inserted.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
                    inserted.HorizontalPositionOffset = 140 * 12700;
                    inserted.VerticalPositionOffset = 160 * 12700;
                }
                if (mixedAnchor) {
                    image.Position = 0;
                    WordImage anchored = p.InsertImage(image, "anchor.png", 64, 32, WordImageTextWrapping.InFrontOfText);
                    anchored.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
                    anchored.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
                    anchored.HorizontalPositionOffset = 140 * 12700;
                    anchored.VerticalPositionOffset = 160 * 12700;
                }
            }
            if (followingBlank) AddBlank();
        }
        document.AddParagraph("B").LineSpacingBeforePoints = neighborSpacing;
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        var letters = pdf.GetPage(1).Letters;
        inspect?.Invoke(pdf.GetPage(1));
        if (mixedAnchor) Assert.Equal(2, pdf.GetPage(1).GetImages().Count());
        return Assert.Single(letters, l => l.Value == "A").StartBaseLine.Y -
            Assert.Single(letters, l => l.Value == "B").StartBaseLine.Y;
    }
}
