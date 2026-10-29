using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests;

public sealed class PowerPointTypedAuthoringLayoutTests {
    public sealed record Revenue(string Period, double Amount);
    private static readonly Revenue[] Rows = { new("Q1", 12D), new("Q2", 17D) };
    private static PowerPointChart Chart(PowerPointSlide slide, OfficeChartOptions? options = null) =>
        slide.AddChart(OfficeChartType.Column, Rows, nameof(Revenue.Period), options, nameof(Revenue.Amount));

    [Theory]
    [InlineData(960D, 540D, false)]
    [InlineData(720D, 540D, false)]
    [InlineData(960D, 540D, true)]
    [InlineData(720D, 540D, true)]
    public void TypedObjectsUseCanonicalContentFramesAndStayEditable(double width, double height, bool titleLast) {
        using var document = PowerPointDocument.Create();
        document.SlideSize.SetSize(OfficeLength.Points(width), OfficeLength.Points(height));
        var slide = document.AddSlide();
        PowerPointTextBox? title = titleLast ? null : slide.AddTitle("Revenue");
        var chart = Chart(slide);
        var nativeChart = chart.Element;
        var table = slide.AddTable(Rows, new OfficeTableOptions());
        title ??= slide.AddTitle("Revenue");
        Assert.Same(nativeChart, chart.Element);
        Assert.Same(chart, slide.Charts.Single());
        Assert.Same(table, slide.Tables.Single());
        Assert.True(chart.Left > OfficeLength.Zero);
        Assert.True(chart.Top > title.Top + title.Height);
        Assert.True(table.Left >= chart.Left + chart.Width);
        Assert.True(table.Top > title.Top + title.Height);
        Assert.True(table.Left + table.Width < document.SlideSize.Width);
        var context = new PowerPointLayoutContext(document, document.LayoutTheme);
        Assert.Equal(OfficeLength.Points(context.Grid.X(1)), chart.Left);
        Assert.Equal(OfficeLength.Points(context.Grid.W(6)), chart.Width);
        Assert.Equal(OfficeLength.Points(context.Grid.X(7)), table.Left);
        table.GetCell(1, 1).Text = "Edited";
        chart.SetTitle("Edited chart");
        using var bytes = new MemoryStream();
        document.Save(bytes);
        bytes.Position = 0;
        using var native = PresentationDocument.Open(bytes, false);
        var part = native.PresentationPart!.SlideParts.Single();
        Assert.Equal(2, part.Slide.Descendants<P.GraphicFrame>().Count());
        Assert.Single(part.ChartParts);
        Assert.Single(part.ChartParts.Single().GetPartsOfType<EmbeddedPackagePart>().Single().GetParentParts());
        Assert.Contains("Edited", part.Slide.Descendants<A.Text>().Select(text => text.Text));
        Assert.Contains("Edited chart", part.ChartParts.Single().ChartSpace.Descendants<A.Text>().Select(text => text.Text));
        Assert.Single(part.ChartParts.Single().GetPartsOfType<EmbeddedPackagePart>());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SingleTypedObjectStartsBelowReservedHeader(bool table) {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        PowerPointShape shape = table ? slide.AddTable(Rows, new OfficeTableOptions()) : Chart(slide);
        Assert.True(shape.Left > OfficeLength.Zero);
        Assert.True(shape.Top > OfficeLength.Zero);
        Assert.Equal(OfficeLength.Points(new PowerPointLayoutContext(document, document.LayoutTheme).Grid.W(12)), shape.Width);
        var title = slide.AddTitle("A title added after content");
        Assert.True(shape.Top > title.Top + title.Height);
    }

    [Fact]
    public void AllExplicitEmusAndSharedChartOptionsRemainExactAfterAnotherAutomaticObject() {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        var fixedBox = new PowerPointLayoutBox(OfficeLength.Emus(127001), OfficeLength.Emus(254003),
            OfficeLength.Emus(1270007), OfficeLength.Emus(1905001));
        var title = slide.AddTitle("Explicit", fixedBox);
        var chart = Chart(slide, new PowerPointChartOptions { Left = fixedBox.Left, Top = fixedBox.Top,
            Width = fixedBox.Width, Height = fixedBox.Height });
        var table = slide.AddTable(Rows, new[] {
            new PowerPointTableColumn<Revenue>("Period", row => row.Period),
            new PowerPointTableColumn<Revenue>("Amount", row => row.Amount)
        }, fixedBox);
        slide.AddTable(2, 2);
        AssertBox(fixedBox, title.Bounds);
        AssertBox(fixedBox, chart.Bounds);
        AssertBox(fixedBox, table.Bounds);
    }

    [Fact]
    public void PartialOptionsPreserveSuppliedLengthsAndUseEngineForMissingComponents() {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        var title = slide.AddTitle("Shared options");
        var chart = Chart(slide, new OfficeChartOptions { Width = OfficeLength.Emus(2540001), Title = "Chart label" });
        Assert.Equal(2540001, chart.Width.ToEmus());
        Assert.True(chart.Left > OfficeLength.Zero);
        Assert.True(chart.Top > title.Top + title.Height);
        var before = chart.Bounds;
        slide.AddTable(Rows, new OfficeTableOptions());
        AssertBox(before, chart.Bounds);
    }

    [Fact]
    public void CallerGeometryEditsEndAutomaticReflowOwnership() {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        var chart = Chart(slide);
        chart.Left = OfficeLength.Emus(17);
        chart.Width = OfficeLength.Emus(2540001);
        var before = chart.Bounds;
        slide.AddTable(2, 2);
        slide.AddTitle("Pinned chart");
        AssertBox(before, chart.Bounds);
    }

    [Fact]
    public void ThirdAutomaticObjectFailsBeforeAddingNativeParts() {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        Chart(slide);
        slide.AddTable(2, 2);
        int shapes = slide.Shapes.Count;
        int parts = slide.SlidePart.Parts.Count();
        var before = slide.Shapes.Select(shape => shape.Bounds).ToArray();
        Assert.Throws<InvalidOperationException>(() => Chart(slide));
        Assert.Equal(shapes, slide.Shapes.Count);
        Assert.Equal(parts, slide.SlidePart.Parts.Count());
        for (int i = 0; i < shapes; i++) AssertBox(before[i], slide.Shapes[i].Bounds);
    }

    [Fact]
    public void AutomaticResizePreservesTableColumnRatios() {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        var table = slide.AddTable(Rows, new OfficeTableOptions());
        table.SetColumnWidthsByRatio(1D, 3D);
        var native = table.Element;
        Chart(slide);
        Assert.Same(native, table.Element);
        Assert.InRange(table.GetColumnWidth(1).ToEmus() / (double)table.GetColumnWidth(0).ToEmus(), 2.999, 3.001);
        Assert.InRange(Enumerable.Range(0, table.Columns).Sum(column => table.GetColumnWidth(column).ToEmus()), table.Width.ToEmus() - 1, table.Width.ToEmus());
    }

    [Fact]
    public void LongAndEmptyTitlesKeepNativeHandlesAndReserveBodySpace() {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        var chart = Chart(slide);
        var title = slide.AddTitle("Quarterly revenue and margin across the regional business units");
        Assert.Same(title, slide.TextBoxes.Single(box => box.PlaceholderType == PowerPointPlaceholderType.Title));
        Assert.True(chart.Top > title.Top + title.Height);
        var emptySlide = document.AddSlide();
        var empty = emptySlide.AddTitle(string.Empty);
        Assert.Same(empty, emptySlide.TextBoxes.Single(box => box.PlaceholderType == PowerPointPlaceholderType.Title));
    }

    [Fact]
    public void TemplateHeaderInheritsNativeStyleAndBodyStaysBelowItsAuthoredFrame() {
        using var template = PowerPointDocument.CreateForImport();
        int layoutIndex = template.GetSlideLayouts().First().LayoutIndex;
        var authored = template.EnsureLayoutPlaceholderTextBox(0, layoutIndex, PowerPointPlaceholderType.Title);
        authored.Text = "Template heading";
        authored.Bounds = new PowerPointLayoutBox(OfficeLength.Points(60), OfficeLength.Points(48),
            OfficeLength.Points(810), OfficeLength.Points(80));
        authored.FontSize = OfficeLength.Points(31);
        using var bytes = new MemoryStream();
        template.Save(bytes);
        bytes.Position = 0;
        using var loaded = PowerPointDocument.Load(bytes);
        var slide = loaded.AddSlide(0, layoutIndex);
        var title = slide.AddTitle("Template title");
        var chart = Chart(slide);
        Assert.Null(((P.Shape)title.Element).ShapeProperties!.GetFirstChild<A.Transform2D>());
        Assert.True(title.TryGetExportBoundsPoints(out double left, out double top, out double width, out double height));
        Assert.Equal(60D, left); Assert.Equal(48D, top); Assert.Equal(810D, width); Assert.Equal(80D, height);
        Assert.True(chart.Top.ToPoints() > top + height);
        Assert.Equal(OfficeLength.Points(31), loaded.GetLayoutPlaceholderTextBox(0, layoutIndex, PowerPointPlaceholderType.Title)!.FontSize);
    }

    [Fact]
    public void NeutralImportKeepsExactImageFrameWhileAuthoringThemeIsAvailable() {
        using var document = PowerPointDocument.CreateForImport();
        using var png = new MemoryStream(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jWZkAAAAASUVORK5CYII="));
        var slide = document.AddSlide(PowerPointSlideLayoutType.Blank);
        var box = new PowerPointLayoutBox(OfficeLength.Emus(0), OfficeLength.Emus(0),
            OfficeLength.Emus(12192001), OfficeLength.Emus(6858001));
        var image = slide.AddImage(png, OfficeIMO.Drawing.OfficeImageFormat.Png, box);
        AssertBox(box, image.Bounds);
        Assert.False(document.OwnsLayoutTheme);
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using var native = PresentationDocument.Open(output, false);
        var frame = native.PresentationPart!.SlideParts.Single().Slide.Descendants<P.Picture>().Single().ShapeProperties!.Transform2D!;
        Assert.Equal(0L, frame.Offset!.X!.Value); Assert.Equal(0L, frame.Offset.Y!.Value);
        Assert.Equal(12192001L, frame.Extents!.Cx!.Value); Assert.Equal(6858001L, frame.Extents.Cy!.Value);
    }

    [Fact]
    public void ExplicitTitleAddedAfterTypedBodyMovesOnlyAutomaticContent() {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        var chart = Chart(slide);
        var title = slide.AddTitle("Positioned title", OfficeLength.Points(60), OfficeLength.Points(190),
            OfficeLength.Points(800), OfficeLength.Points(55));
        Assert.Equal(OfficeLength.Points(190), title.Top);
        Assert.Equal(OfficeLength.Points(55), title.Height);
        Assert.True(chart.Top > title.Top + title.Height);
        Assert.Same(chart, slide.Charts.Single());
    }

    private static void AssertBox(PowerPointLayoutBox expected, PowerPointLayoutBox actual) {
        Assert.Equal(expected.Left, actual.Left); Assert.Equal(expected.Top, actual.Top);
        Assert.Equal(expected.Width, actual.Width); Assert.Equal(expected.Height, actual.Height);
    }
}
