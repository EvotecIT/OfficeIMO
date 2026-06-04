using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void Paragraph_RendersExplicitLineBreaksInsideRichText() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 12
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Bold("Finding").LineBreak().Text("No critical issues detected."))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double findingY = FindWordStartY(page, "Finding");
        double noY = FindWordStartY(page, "No");

        Assert.True(findingY > noY + 12, $"Expected the explicit line break to move following text to the next line. Finding y: {findingY:0.##}, No y: {noY:0.##}.");
    }

    [Fact]
    public void Paragraph_UsesConfiguredSpacingBeforeAndAfter() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] defaultBytes = CreateParagraphSpacingProbe(options, null);
        byte[] spacedBytes = CreateParagraphSpacingProbe(options, new PdfParagraphStyle {
            SpacingBefore = 12,
            SpacingAfter = 18
        });

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var spacedPage = spacedPdf.GetPage(1);

        double defaultTargetY = FindWordStartY(defaultPage, "TargetMarker");
        double spacedTargetY = FindWordStartY(spacedPage, "TargetMarker");
        double defaultAfterY = FindWordStartY(defaultPage, "AfterMarker");
        double spacedAfterY = FindWordStartY(spacedPage, "AfterMarker");

        Assert.True(defaultTargetY - spacedTargetY >= 10, $"Expected paragraph spacing before to move target text down. Default y: {defaultTargetY:0.##}, spaced y: {spacedTargetY:0.##}.");
        Assert.True(defaultAfterY - spacedAfterY >= 24, $"Expected paragraph spacing before and after to move following text down. Default y: {defaultAfterY:0.##}, spaced y: {spacedAfterY:0.##}.");
    }

    [Fact]
    public void Paragraph_SuppressesSpacingBeforeAtPageTop() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] defaultBytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("TopMarker"))
            .ToBytes();
        byte[] spacedBytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("TopMarker"), style: new PdfParagraphStyle {
                SpacingBefore = 28,
                SpacingAfter = 0
            })
            .ToBytes();

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(1), "TopMarker");
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(1), "TopMarker");

        Assert.InRange(Math.Abs(defaultTopY - spacedTopY), 0, 1.5);
    }

    [Fact]
    public void Paragraph_SuppressesSpacingBeforeAtRowColumnTop() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] defaultBytes = PdfDocument.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column
                .Paragraph(p => p.Text("ColumnTopMarker")))))))
            .ToBytes();
        byte[] spacedBytes = PdfDocument.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column
                .Paragraph(p => p.Text("ColumnTopMarker"), style: new PdfParagraphStyle {
                    SpacingBefore = 28,
                    SpacingAfter = 0
                }))))))
            .ToBytes();

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(1), "ColumnTopMarker");
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(1), "ColumnTopMarker");

        Assert.InRange(Math.Abs(defaultTopY - spacedTopY), 0, 1.5);
    }

    [Fact]
    public void Spacer_AddsInvisibleVerticalSpaceWithoutExtractedText() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        var paragraphStyle = new PdfParagraphStyle { SpacingBefore = 0, SpacingAfter = 0 };
        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("BeforeMarker"), style: paragraphStyle)
            .Spacer(24)
            .Paragraph(p => p.Text("AfterMarker"), style: paragraphStyle)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        Assert.Equal(2, CountTextLines(page));
        Assert.Contains("BeforeMarker", page.Text, StringComparison.Ordinal);
        Assert.Contains("AfterMarker", page.Text, StringComparison.Ordinal);

        double beforeY = FindWordStartY(page, "BeforeMarker");
        double afterY = FindWordStartY(page, "AfterMarker");
        Assert.InRange(beforeY - afterY, 36, 42);
    }

    [Fact]
    public void Spacer_WorksInsideRowColumnFlow() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        var paragraphStyle = new PdfParagraphStyle { SpacingBefore = 0, SpacingAfter = 0 };
        byte[] bytes = PdfDocument.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column
                .Paragraph(p => p.Text("TopMarker"), style: paragraphStyle)
                .Spacer(20)
                .Paragraph(p => p.Text("BottomMarker"), style: paragraphStyle))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        Assert.Equal(2, CountTextLines(page));

        double topY = FindWordStartY(page, "TopMarker");
        double bottomY = FindWordStartY(page, "BottomMarker");
        Assert.InRange(topY - bottomY, 32, 38);
    }

    [Fact]
    public void Spacer_RejectsInvalidHeights() {
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().Spacer(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().Spacer(double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().Spacer(double.PositiveInfinity));
    }

    [Fact]
    public void Paragraph_UsesConfiguredLineHeight() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] defaultBytes = CreateParagraphLineHeightProbe(options, null);
        byte[] looseBytes = CreateParagraphLineHeightProbe(options, new PdfParagraphStyle {
            LineHeight = 2.0,
            SpacingAfter = 0
        });

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var loosePdf = PdfPigDocument.Open(new MemoryStream(looseBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var loosePage = loosePdf.GetPage(1);

        double defaultFirstY = FindWordStartY(defaultPage, "FirstLine");
        double defaultSecondY = FindWordStartY(defaultPage, "SecondLine");
        double defaultThirdY = FindWordStartY(defaultPage, "ThirdLine");
        double looseFirstY = FindWordStartY(loosePage, "FirstLine");
        double looseSecondY = FindWordStartY(loosePage, "SecondLine");
        double looseThirdY = FindWordStartY(loosePage, "ThirdLine");

        double defaultGapOne = defaultFirstY - defaultSecondY;
        double defaultGapTwo = defaultSecondY - defaultThirdY;
        double looseGapOne = looseFirstY - looseSecondY;
        double looseGapTwo = looseSecondY - looseThirdY;

        Assert.True(looseGapOne - defaultGapOne >= 5, $"Expected configured line height to increase the first line gap. Default gap: {defaultGapOne:0.##}, loose gap: {looseGapOne:0.##}.");
        Assert.True(looseGapTwo - defaultGapTwo >= 5, $"Expected configured line height to increase the second line gap. Default gap: {defaultGapTwo:0.##}, loose gap: {looseGapTwo:0.##}.");
    }

    [Fact]
    public void Paragraph_UsesConfiguredHorizontalIndents() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 280,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] defaultBytes = CreateParagraphIndentProbe(options, null);
        byte[] indentedBytes = CreateParagraphIndentProbe(options, new PdfParagraphStyle {
            LeftIndent = 24,
            RightIndent = 90,
            SpacingAfter = 0
        });

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var indentedPdf = PdfPigDocument.Open(new MemoryStream(indentedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var indentedPage = indentedPdf.GetPage(1);

        double defaultX = FindWordStartX(defaultPage, "IndentedMarker");
        double indentedX = FindWordStartX(indentedPage, "IndentedMarker");
        int defaultLineCount = CountTextLines(defaultPage);
        int indentedLineCount = CountTextLines(indentedPage);

        Assert.True(indentedX - defaultX >= 22, $"Expected left indent to move paragraph text right. Default x: {defaultX:0.##}, indented x: {indentedX:0.##}.");
        Assert.True(indentedLineCount > defaultLineCount, $"Expected right indent to reduce text width and increase wrapping. Default lines: {defaultLineCount}, indented lines: {indentedLineCount}.");
    }

    [Fact]
    public void Paragraph_UsesConfiguredFirstLineIndent() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p
                .Text("FirstIndentMarker")
                .LineBreak()
                .Text("SecondIndentMarker"), style: new PdfParagraphStyle {
                    FirstLineIndent = 24,
                    SpacingAfter = 0
                })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double firstX = FindWordStartX(page, "FirstIndentMarker");
        double secondX = FindWordStartX(page, "SecondIndentMarker");

        Assert.True(firstX - secondX >= 22, $"Expected first line indent to move only the first line right. First x: {firstX:0.##}, second x: {secondX:0.##}.");
    }

    [Fact]
    public void Paragraph_UsesDefaultParagraphStyleWhenStyleIsNotProvided() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10,
            DefaultParagraphStyle = new PdfParagraphStyle {
                FirstLineIndent = 24,
                SpacingAfter = 0
            }
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p
                .Text("DefaultFirstIndent")
                .LineBreak()
                .Text("DefaultSecondIndent"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double firstX = FindWordStartX(page, "DefaultFirstIndent");
        double secondX = FindWordStartX(page, "DefaultSecondIndent");

        Assert.True(firstX - secondX >= 22, $"Expected default paragraph style to indent only the first line. First x: {firstX:0.##}, second x: {secondX:0.##}.");
    }

    [Fact]
    public void PdfDocument_DefaultParagraphStyleAppliesToFollowingParagraphsAndSnapshotsInput() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = new PdfParagraphStyle {
            FirstLineIndent = 24,
            SpacingAfter = 0
        };

        byte[] bytes = PdfDocument.Create(options)
            .DefaultParagraphStyle(style)
            .Paragraph(p => p
                .Text("FluentDefaultFirst")
                .LineBreak()
                .Text("FluentDefaultSecond"))
            .ToBytes();

        style.FirstLineIndent = 0;

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double firstX = FindWordStartX(page, "FluentDefaultFirst");
        double secondX = FindWordStartX(page, "FluentDefaultSecond");

        Assert.True(firstX - secondX >= 22, $"Expected fluent default paragraph style to indent only the first line and snapshot caller input. First x: {firstX:0.##}, second x: {secondX:0.##}.");
    }

    [Fact]
    public void PdfDocument_DefaultParagraphStyleRejectsNull() {
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultParagraphStyle(null!));
    }

    [Fact]
    public void Paragraph_ExplicitStyleOverridesDefaultParagraphStyle() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10,
            DefaultParagraphStyle = new PdfParagraphStyle {
                LeftIndent = 40,
                SpacingAfter = 0
            }
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("ExplicitMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 0
            })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double markerX = FindWordStartX(page, "ExplicitMarker");

        Assert.InRange(markerX, options.MarginLeft - 1, options.MarginLeft + 3);
    }

    [Fact]
    public void Paragraph_UsesConfiguredHangingIndent() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p
                .Text("HangingFirst")
                .LineBreak()
                .Text("HangingSecond"), style: new PdfParagraphStyle {
                    LeftIndent = 24,
                    FirstLineIndent = -24,
                    SpacingAfter = 0
                })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double firstX = FindWordStartX(page, "HangingFirst");
        double secondX = FindWordStartX(page, "HangingSecond");

        Assert.True(secondX - firstX >= 22, $"Expected hanging indent to move following lines right of the first line. First x: {firstX:0.##}, second x: {secondX:0.##}.");
    }

    [Fact]
    public void RowColumnParagraph_UsesConfiguredFirstLineIndent() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Paragraph(p => p
                                    .Text("ColumnFirstIndent")
                                    .LineBreak()
                                    .Text("ColumnSecondIndent"), style: new PdfParagraphStyle {
                                        FirstLineIndent = 24,
                                        SpacingAfter = 0
                                    }))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double firstX = FindWordStartX(page, "ColumnFirstIndent");
        double secondX = FindWordStartX(page, "ColumnSecondIndent");

        Assert.True(firstX - secondX >= 22, $"Expected row column first line indent to move only the first line right. First x: {firstX:0.##}, second x: {secondX:0.##}.");
    }

    [Fact]
    public void RowColumnParagraph_UsesDefaultParagraphStyleWhenStyleIsNotProvided() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10,
            DefaultParagraphStyle = new PdfParagraphStyle {
                FirstLineIndent = 24,
                SpacingAfter = 0
            }
        };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Paragraph(p => p
                                    .Text("ColumnDefaultFirst")
                                    .LineBreak()
                                    .Text("ColumnDefaultSecond")))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double firstX = FindWordStartX(page, "ColumnDefaultFirst");
        double secondX = FindWordStartX(page, "ColumnDefaultSecond");

        Assert.True(firstX - secondX >= 22, $"Expected row column default paragraph style to indent only the first line. First x: {firstX:0.##}, second x: {secondX:0.##}.");
    }


}
