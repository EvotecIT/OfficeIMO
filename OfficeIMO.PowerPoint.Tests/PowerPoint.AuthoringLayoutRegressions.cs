using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests;

public sealed class PowerPointAuthoringLayoutRegressions {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public void NativeTitleIdentityAndGeometrySurviveEmptyAndInheritedVariants(bool centered, bool inherited, bool empty) {
        using var template = PowerPointDocument.CreateForImport();
        int layoutIndex = template.GetSlideLayouts().First().LayoutIndex;
        PowerPointPlaceholderType type = centered ? PowerPointPlaceholderType.CenteredTitle : PowerPointPlaceholderType.Title;
        var layoutTree = template.GetSlideLayoutPart(0, layoutIndex).SlideLayout.CommonSlideData!.ShapeTree!;
        foreach (var old in layoutTree.Elements<P.Shape>().Where(shape =>
            shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value is { } value &&
            (value == P.PlaceholderValues.Title || value == P.PlaceholderValues.CenteredTitle)).ToArray()) old.Remove();
        var native = template.EnsureLayoutPlaceholderTextBox(0, layoutIndex, type, index: 7U);
        native.Text = "Native title";
        native.Bounds = new PowerPointLayoutBox(OfficeLength.Points(60), OfficeLength.Points(48), OfficeLength.Points(600), OfficeLength.Points(80));
        native.FontSize = OfficeLength.Points(31);
        native.Color = "123456";
        var layoutShape = (P.Shape)native.Element;
        if (inherited) {
            var master = template.OpenXmlDocument.PresentationPart!.SlideMasterParts.First().SlideMaster;
            foreach (var old in master.CommonSlideData!.ShapeTree!.Elements<P.Shape>().Where(shape =>
                shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value is { } value &&
                (value == P.PlaceholderValues.Title || value == P.PlaceholderValues.CenteredTitle)).ToArray()) old.Remove();
            var masterShape = (P.Shape)layoutShape.CloneNode(true);
            masterShape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id = 42000U;
            master.CommonSlideData.ShapeTree.Append(masterShape);
            layoutShape.ShapeProperties!.Transform2D?.Remove();
        }
        using var bytes = new MemoryStream();
        template.Save(bytes); bytes.Position = 0;
        using var document = PowerPointDocument.Load(bytes);
        var slide = document.AddSlide(0, layoutIndex);
        var title = slide.AddTitle(empty ? string.Empty : "Inserted title");
        var shape = (P.Shape)title.Element;
        var placeholder = shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.PlaceholderShape!;
        Assert.Equal(type, title.PlaceholderType);
        Assert.Equal(7U, placeholder.Index!.Value);
        Assert.Null(shape.ShapeProperties!.Transform2D);
        Assert.True(title.TryGetExportBoundsPoints(out double left, out double top, out double width, out double height));
        Assert.Equal(60D, left); Assert.Equal(48D, top); Assert.Equal(600D, width); Assert.Equal(80D, height);
        Assert.All(shape.Descendants<A.RunProperties>(), run => { Assert.Null(run.FontSize); Assert.Empty(run.Elements<A.SolidFill>()); });
        Assert.All(shape.Descendants<A.EndParagraphRunProperties>(), end => { Assert.Null(end.FontSize); Assert.Empty(end.GetAttributes()); });
        Assert.Equal(empty ? string.Empty : "Inserted title", title.Text);
        var chart = slide.AddChart(OfficeChartType.Column, new OfficeIMO.Drawing.OfficeChartData(new[] { "Q1" },
            new[] { new OfficeIMO.OfficeSeries("Revenue", new[] { 12D }) }));
        Assert.True(chart.Top.ToPoints() > top + height);
        using var saved = new MemoryStream(); document.Save(saved);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReloadedStandardSizeUsesAnInBoundsAutomaticFrame(bool table) {
        using var source = PowerPointDocument.Create();
        source.SlideSize.SetSize(OfficeLength.Points(720), OfficeLength.Points(540));
        source.AddSlide().AddTitle("Standard size");
        using var bytes = new MemoryStream(); source.Save(bytes); bytes.Position = 0;
        using var document = PowerPointDocument.Load(bytes);
        var slide = document.AddSlide();
        var title = slide.AddTitle("After reload");
        PowerPointShape body = table ? slide.AddTable(new[] { new { Quarter = "Q1", Revenue = 12 } })
            : slide.AddChart(OfficeChartType.Column, new OfficeIMO.Drawing.OfficeChartData(new[] { "Q1" },
                new[] { new OfficeIMO.OfficeSeries("Revenue", new[] { 12D }) }));
        Assert.True(body.Left >= OfficeLength.Zero);
        Assert.True(body.Top >= OfficeLength.Zero);
        Assert.True(body.Left + body.Width <= document.SlideSize.Width);
        Assert.True(body.Top + body.Height <= document.SlideSize.Height);
        Assert.True(title.TryGetExportBoundsPoints(out double x, out _, out double width, out _));
        Assert.True(x + width <= document.SlideSize.Width.ToPoints());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FailedExplicitTitleDoesNotAttachOrMoveContent(bool partial) {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        var chart = slide.AddChart(OfficeChartType.Column, new OfficeIMO.Drawing.OfficeChartData(new[] { "Q1" },
            new[] { new OfficeIMO.OfficeSeries("Revenue", new[] { 12D }) }));
        var original = chart.Bounds;
        int shapes = slide.Shapes.Count;
        Assert.Throws<InvalidOperationException>(() => slide.AddTitle("No content space", partial ? null : OfficeLength.Points(48),
            OfficeLength.Points(480), partial ? null : OfficeLength.Points(600), OfficeLength.Points(24)));
        Assert.Equal(shapes, slide.Shapes.Count);
        Assert.Empty(slide.TextBoxes);
        Assert.Equal(original.Left, chart.Left); Assert.Equal(original.Top, chart.Top);
        Assert.Equal(original.Width, chart.Width); Assert.Equal(original.Height, chart.Height);
        Assert.NotNull(slide.AddTitle("Still usable"));
    }

    [Fact]
    public void DarkDefaultTableKeepsReadableBodyColorWithoutDirectOverrides() {
        using var document = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = OfficeTheme.Dark });
        var table = document.AddSlide().AddTable(new[] { new { Quarter = "Q1", Revenue = 12 } });
        var style = document.OpenXmlDocument.PresentationPart!.TableStylesPart!.TableStyleList!.Elements<A.TableStyleEntry>()
            .Single(item => item.StyleId!.Value == "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
        var color = style.GetFirstChild<A.WholeTable>()!.GetFirstChild<A.TableCellTextStyle>()!.GetFirstChild<A.RgbColorModelHex>();
        Assert.Equal(OfficeTheme.Dark.Palette.Background.ToRgbHex(), color!.Val!.Value);
        Assert.Empty(table.GetCell(1, 0).Cell.Descendants<A.RunProperties>().SelectMany(run => run.Elements<A.SolidFill>()));
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        Assert.Single(reopened.Slides.Single().Tables);
    }
}
