using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointSharedChartTextStyleTests {
    [Fact]
    public void SharedSnapshot_PreservesExplicitNativeChartFonts() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        var data = new OfficeChartData(
            new[] { "Q1", "Q2" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 20D }) });
        PowerPointChart chart = slide.AddChartPoints(
            OfficeChartKind.ColumnClustered, data, 40, 40, 500, 300);
        chart.SetTitle("Trajectory")
            .SetTitleTextStyle(fontName: "Arial")
            .SetLegendTextStyle(fontName: "Arial")
            .SetCategoryAxisLabelTextStyle(fontName: "Arial")
            .SetValueAxisLabelTextStyle(fontName: "Arial");

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Arial", snapshot.Style.FontFamily);
        Assert.Equal("Arial", snapshot.Style.TitleFontFamily);
    }

    [Fact]
    public void SharedSnapshot_DoesNotProjectTitleFontOntoChartBodyText() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        var data = new OfficeChartData(
            new[] { "Q1", "Q2" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 20D }) });
        PowerPointChart chart = slide.AddChartPoints(
            OfficeChartKind.ColumnClustered, data, 40, 40, 500, 300);
        chart.SetTitle("Trajectory")
            .SetTitleTextStyle(fontName: "Georgia")
            .SetLegendTextStyle(fontName: "Arial")
            .SetCategoryAxisLabelTextStyle(fontName: "Arial")
            .SetValueAxisLabelTextStyle(fontName: "Arial");

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Arial", snapshot.Style.FontFamily);
        Assert.Equal("Georgia", snapshot.Style.TitleFontFamily);
    }

    [Fact]
    public void SharedSnapshot_DoesNotFlattenMixedTitleRunFonts() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointChart chart = presentation.AddSlide().AddChartPoints(
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "Q1" }, new[] {
                new OfficeChartSeries("Actual", new[] { 10D })
            }), 40, 40, 500, 300);
        chart.SetTitle("First").SetTitleTextStyle(fontName: "Arial");
        C.Chart openXmlChart = presentation.OpenXmlDocument.PresentationPart!
            .SlideParts.Single().ChartParts.Single().ChartSpace!
            .GetFirstChild<C.Chart>()!;
        A.Paragraph paragraph = openXmlChart.GetFirstChild<C.Title>()!
            .GetFirstChild<C.ChartText>()!.GetFirstChild<C.RichText>()!
            .GetFirstChild<A.Paragraph>()!;
        paragraph.Append(new A.Run(
            new A.RunProperties(new A.LatinFont { Typeface = "Georgia" }),
            new A.Text("Second")));

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(OfficeChartStyle.Default.TitleFontFamily,
            snapshot.Style.TitleFontFamily);
    }

    [Fact]
    public void SharedSnapshot_DoesNotFlattenMixedTitleFieldFonts() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointChart chart = presentation.AddSlide().AddChartPoints(
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "Q1" }, new[] {
                new OfficeChartSeries("Actual", new[] { 10D })
            }), 40, 40, 500, 300);
        chart.SetTitle("First").SetTitleTextStyle(fontName: "Arial");
        C.Chart openXmlChart = presentation.OpenXmlDocument.PresentationPart!
            .SlideParts.Single().ChartParts.Single().ChartSpace!
            .GetFirstChild<C.Chart>()!;
        A.Paragraph paragraph = openXmlChart.GetFirstChild<C.Title>()!
            .GetFirstChild<C.ChartText>()!.GetFirstChild<C.RichText>()!
            .GetFirstChild<A.Paragraph>()!;
        paragraph.Append(new A.Field(
            new A.RunProperties(new A.LatinFont { Typeface = "Georgia" }),
            new A.Text("Second")) {
            Id = "{00000000-0000-0000-0000-000000000001}",
            Type = "datetime"
        });

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(OfficeChartStyle.Default.TitleFontFamily,
            snapshot.Style.TitleFontFamily);
    }

    [Fact]
    public void SharedSnapshot_UsesTitleParagraphDefaultRunFont() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointChart chart = presentation.AddSlide().AddChartPoints(
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "Q1" }, new[] {
                new OfficeChartSeries("Actual", new[] { 10D })
            }), 40, 40, 500, 300);
        chart.SetTitle("Trajectory");
        C.Chart openXmlChart = presentation.OpenXmlDocument.PresentationPart!
            .SlideParts.Single().ChartParts.Single().ChartSpace!
            .GetFirstChild<C.Chart>()!;
        A.Paragraph paragraph = openXmlChart.GetFirstChild<C.Title>()!
            .GetFirstChild<C.ChartText>()!.GetFirstChild<C.RichText>()!
            .GetFirstChild<A.Paragraph>()!;
        foreach (A.RunProperties properties in paragraph
                     .Descendants<A.RunProperties>()) {
            properties.RemoveAllChildren<A.LatinFont>();
        }
        paragraph.ParagraphProperties ??= new A.ParagraphProperties();
        paragraph.ParagraphProperties.RemoveAllChildren<A.DefaultRunProperties>();
        paragraph.ParagraphProperties.Append(new A.DefaultRunProperties(
            new A.LatinFont { Typeface = "Georgia" }));

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Georgia", snapshot.Style.TitleFontFamily);
    }

    [Fact]
    public void SharedSnapshot_UsesTitleListLevelDefaultRunFont() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointChart chart = presentation.AddSlide().AddChartPoints(
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "Q1" }, new[] {
                new OfficeChartSeries("Actual", new[] { 10D })
            }), 40, 40, 500, 300);
        chart.SetTitle("Trajectory");
        C.RichText richText = presentation.OpenXmlDocument.PresentationPart!
            .SlideParts.Single().ChartParts.Single().ChartSpace!
            .GetFirstChild<C.Chart>()!.GetFirstChild<C.Title>()!
            .GetFirstChild<C.ChartText>()!.GetFirstChild<C.RichText>()!;
        A.Paragraph paragraph = richText.GetFirstChild<A.Paragraph>()!;
        foreach (A.RunProperties properties in paragraph
                     .Descendants<A.RunProperties>()) {
            properties.RemoveAllChildren<A.LatinFont>();
        }
        paragraph.ParagraphProperties?.RemoveAllChildren<A.DefaultRunProperties>();
        A.ListStyle listStyle = richText.GetFirstChild<A.ListStyle>()!;
        listStyle.RemoveAllChildren();
        listStyle.Append(new A.Level1ParagraphProperties(
            new A.DefaultRunProperties(
                new A.LatinFont { Typeface = "Georgia" })));

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Georgia", snapshot.Style.TitleFontFamily);
    }

    [Fact]
    public void SharedSnapshot_DoesNotPromoteConflictingBodyFonts() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        var data = new OfficeChartData(
            new[] { "Q1", "Q2" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 20D }) });
        PowerPointChart chart = slide.AddChartPoints(
            OfficeChartKind.ColumnClustered, data, 40, 40, 500, 300);
        chart.SetLegendTextStyle(fontName: "Arial")
            .SetCategoryAxisLabelTextStyle(fontName: "Georgia")
            .SetValueAxisLabelTextStyle(fontName: "Arial");

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.NotNull(snapshot.Style);
        Assert.Equal(OfficeChartStyle.Default.FontFamily,
            snapshot.Style.FontFamily);
    }

    [Fact]
    public void SharedSnapshot_DoesNotPromoteOnePartialBodyFont() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        var data = new OfficeChartData(
            new[] { "Q1", "Q2" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 20D }) });
        PowerPointChart chart = slide.AddChartPoints(
            OfficeChartKind.ColumnClustered, data, 40, 40, 500, 300);
        chart.SetLegendTextStyle(fontName: "Arial");

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(OfficeChartStyle.Default.FontFamily,
            snapshot.Style.FontFamily);
    }

    [Fact]
    public void SharedSnapshot_ResolvesThemeFontTokensToEffectiveFamilies() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.SetThemeLatinFonts("Theme Heading", "Theme Body");
        PowerPointSlide slide = presentation.AddSlide();
        var data = new OfficeChartData(
            new[] { "Q1", "Q2" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 20D }) });
        PowerPointChart chart = slide.AddChartPoints(
            OfficeChartKind.ColumnClustered, data, 40, 40, 500, 300);
        chart.SetTitle("Trajectory")
            .SetTitleTextStyle(fontName: "+mj-lt")
            .SetLegendTextStyle(fontName: "+mn-lt")
            .SetCategoryAxisLabelTextStyle(fontName: "+mn-lt")
            .SetValueAxisLabelTextStyle(fontName: "+mn-lt");

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Theme Body", snapshot.Style.FontFamily);
        Assert.Equal("Theme Heading", snapshot.Style.TitleFontFamily);
    }

    [Fact]
    public void SharedSnapshot_ResolvesInheritedBodyFontFromMinorTheme() {
        using PowerPointPresentation presentation =
            PowerPointPresentation.Create();
        presentation.SetThemeLatinFonts("Theme Heading", "Theme Body");
        PowerPointChart chart = presentation.AddSlide().AddChartPoints(
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "Q1", "Q2" }, new[] {
                new OfficeChartSeries("Actual", new[] { 10D, 20D })
            }), 40, 40, 500, 300);
        chart.SetLegendTextStyle(fontName: "Arial")
            .SetCategoryAxisLabelTextStyle(fontName: "Arial")
            .SetValueAxisLabelTextStyle(fontName: "Arial");
        C.Chart openXmlChart = presentation.OpenXmlDocument.PresentationPart!
            .SlideParts.Single().ChartParts.Single().ChartSpace!
            .GetFirstChild<C.Chart>()!;
        foreach (A.LatinFont latin in openXmlChart.Descendants<A.LatinFont>()) {
            latin.Remove();
        }

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Theme Body", snapshot.Style.FontFamily);
    }

    [Fact]
    public void SharedSnapshot_UsesChartSpaceDefaultForInheritedBodyText() {
        using PowerPointPresentation presentation =
            PowerPointPresentation.Create();
        PowerPointChart chart = presentation.AddSlide().AddChartPoints(
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "Q1", "Q2" }, new[] {
                new OfficeChartSeries("Actual", new[] { 10D, 20D })
            }), 40, 40, 500, 300);
        C.ChartSpace chartSpace = presentation.OpenXmlDocument
            .PresentationPart!.SlideParts.Single().ChartParts.Single()
            .ChartSpace!;
        foreach (A.LatinFont latin in chartSpace.GetFirstChild<C.Chart>()!
                     .Descendants<A.LatinFont>()) {
            latin.Remove();
        }
        chartSpace.GetFirstChild<C.TextProperties>()?.Remove();
        chartSpace.Append(new C.TextProperties(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph(
                new A.ParagraphProperties(
                    new A.DefaultRunProperties(
                        new A.LatinFont { Typeface = "Chart Default" })),
                new A.EndParagraphRunProperties())));

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Chart Default", snapshot.Style.FontFamily);
    }

    [Fact]
    public void SharedSnapshot_PreservesIndependentlyInheritedTitleFont() {
        using PowerPointPresentation presentation =
            PowerPointPresentation.Create();
        presentation.SetThemeLatinFonts("Theme Heading", "Theme Body");
        PowerPointChart chart = presentation.AddSlide().AddChartPoints(
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "Q1", "Q2" }, new[] {
                new OfficeChartSeries("Actual", new[] { 10D, 20D })
            }), 40, 40, 500, 300);
        chart.SetTitle("Trajectory")
            .SetLegendTextStyle(fontName: "Arial")
            .SetCategoryAxisLabelTextStyle(fontName: "Arial")
            .SetValueAxisLabelTextStyle(fontName: "Arial");
        C.Title title = presentation.OpenXmlDocument.PresentationPart!
            .SlideParts.Single().ChartParts.Single().ChartSpace!
            .GetFirstChild<C.Chart>()!.GetFirstChild<C.Title>()!;
        foreach (A.LatinFont latin in title.Descendants<A.LatinFont>()) {
            latin.Remove();
        }

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Arial", snapshot.Style.FontFamily);
        Assert.Equal("Theme Heading", snapshot.Style.TitleFontFamily);
    }
}
