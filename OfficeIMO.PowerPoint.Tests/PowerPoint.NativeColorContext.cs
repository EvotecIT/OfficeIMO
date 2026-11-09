using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests;

public sealed class PowerPointNativeColorContextTests {
    [Theory]
    [InlineData("Success", "123456", 50196)]
    [InlineData("Warning", "654321", 37647)]
    [InlineData("Danger", "ABCDEF", 25098)]
    public void StatusPaintUsesTheConsumingPaletteAcrossNativeAuthors(string roleName, string rgb, int alpha) {
        OfficeColor role = OfficeColor.Parse(roleName);
        var theme = OfficeTheme.Default.WithPalette(OfficeTheme.Default.Palette.With(
            success: OfficeColor.Parse("12345680"), warning: OfficeColor.Parse("65432160"), danger: OfficeColor.Parse("ABCDEF40")));
        using var document = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = theme });
        var slide = document.AddSlide();
        var cell = slide.AddTable(1, 1).GetCell(0, 0);
        cell.Text = "Cell";
        cell.FillColor = role;
        cell.BorderColor = role;
        cell.SetColor(role);
        cell.Runs[0].BackgroundColor = role;
        var box = slide.AddTextBox("Text");
        box.FillColor = role;
        box.OutlineColor = role;
        box.SetColor(role);
        box.Paragraphs[0].SetColor(role);
        box.ApplyTextStyle(new PowerPointTextStyle(color: role, backgroundColor: role), applyToRuns: true);
        box.SetShadow(role, transparencyPercent: 0);
        box.SetGlow(role, transparencyPercent: 0);
        slide.SetBackgroundColor(role);
        var gradientSlide = document.AddSlide();
        gradientSlide.SetBackgroundGradient(role, OfficeColor.Accent3);
        var chart = slide.AddChart(OfficeChartType.Line, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeSeries("Values", new[] { 1d, 2d }) }));
        chart.SetTitle("Status").SetTitleTextStyle(color: role).SetLegendTextStyle(color: role)
            .SetChartAreaStyle(fillColor: role, lineColor: role).SetPlotAreaStyle(fillColor: role, lineColor: role)
            .SetSeriesFillColor(0, role).SetSeriesLineColor(0, role)
            .SetSeriesMarker(0, OfficeChartMarkerShape.Circle, fillColor: role, lineColor: role)
            .SetCategoryAxisLabelTextStyle(color: role).SetValueAxisLabelTextStyle(color: role)
            .SetValueAxisGridlines(lineColor: role).SetSeriesTrendline(0, OfficeChartTrendlineType.Linear, lineColor: role)
            .SetDataLabelTextStyle(color: role).SetDataLabelShapeStyle(fillColor: role, lineColor: role)
            .SetSeriesDataLabelTemplate(0, new PowerPointChartDataLabelTemplate { TextColor = role, FillColor = role, LineColor = role });
        document.SetLayoutPlaceholderTextStyle(0, 0, PowerPointPlaceholderType.Body,
            new PowerPointTextStyle(color: role, backgroundColor: role), createIfMissing: true);
        using var package = PresentationDocument.Open(new MemoryStream(document.ToBytes()), false);
        var part = package.PresentationPart!.SlideParts.First();
        var nativeCell = part.Slide.Descendants<A.TableCell>().Single();
        AssertPaint(nativeCell.TableCellProperties!, rgb, alpha);
        AssertPaint(nativeCell.TableCellProperties!.LeftBorderLineProperties!, rgb, alpha);
        AssertPaint(nativeCell.TextBody!.Descendants<A.RunProperties>().First(), rgb, alpha);
        AssertColor(nativeCell.TextBody.Descendants<A.Highlight>().First().FirstChild!, rgb, alpha);
        var nativeBox = part.Slide.Descendants<P.Shape>().Single();
        AssertPaint(nativeBox.ShapeProperties!, rgb, alpha);
        AssertPaint(nativeBox.ShapeProperties!.GetFirstChild<A.Outline>()!, rgb, alpha);
        AssertPaint(nativeBox.TextBody!.Descendants<A.RunProperties>().First(), rgb, alpha);
        AssertColor(nativeBox.ShapeProperties.Descendants<A.OuterShadow>().Single().FirstChild!, rgb, alpha);
        AssertColor(nativeBox.ShapeProperties.Descendants<A.Glow>().Single().FirstChild!, rgb, alpha);
        AssertPaint(part.Slide.CommonSlideData!.Background!.BackgroundProperties!, rgb, alpha);
        var gradient = package.PresentationPart.SlideParts.Skip(1).Single().Slide.Descendants<A.GradientStop>().ToArray();
        AssertColor(gradient[0].FirstChild!, rgb, alpha);
        Assert.Equal("accent3", gradient[1].GetFirstChild<A.SchemeColor>()!.Val!.InnerText);
        C.ChartSpace space = part.ChartParts.Single().ChartSpace;
        AssertPaint(space.Descendants<C.Title>().Single().Descendants<A.RunProperties>().First(), rgb, alpha);
        AssertPaint(space.Descendants<C.Legend>().Single().Descendants<A.DefaultRunProperties>().First(), rgb, alpha);
        AssertPaint(space.GetFirstChild<C.ShapeProperties>()!, rgb, alpha);
        AssertPaint(space.Descendants<C.PlotArea>().Single().GetFirstChild<C.ShapeProperties>()!, rgb, alpha);
        var series = space.Descendants<C.LineChartSeries>().Single();
        AssertPaint(series.GetFirstChild<C.ChartShapeProperties>()!, rgb, alpha);
        AssertPaint(series.GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!, rgb, alpha);
        AssertPaint(series.GetFirstChild<C.Marker>()!.GetFirstChild<C.ChartShapeProperties>()!, rgb, alpha);
        AssertPaint(series.GetFirstChild<C.Trendline>()!.GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!, rgb, alpha);
        Assert.All(space.Descendants<C.DataLabels>(), labels => {
            AssertPaint(labels.GetFirstChild<C.ChartShapeProperties>()!, rgb, alpha);
            AssertPaint(labels.Descendants<A.DefaultRunProperties>().First(), rgb, alpha);
        });
        foreach (var axis in space.Descendants<OpenXmlCompositeElement>().Where(element => element is C.CategoryAxis or C.ValueAxis)) {
            AssertPaint(axis.GetFirstChild<C.TextProperties>()!.Descendants<A.DefaultRunProperties>().First(), rgb, alpha);
        }
        AssertPaint(space.Descendants<C.MajorGridlines>().Single().GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!, rgb, alpha);
        var placeholder = package.PresentationPart.SlideMasterParts.First().SlideLayoutParts.First()
            .SlideLayout.Descendants<P.Shape>().Single(shape => shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>()?.Type?.Value == P.PlaceholderValues.Body);
        AssertPaint(placeholder.TextBody!.Descendants<A.DefaultRunProperties>().First(), rgb, alpha);
        using var reopened = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(OfficeColor.Parse(rgb + (roleName == "Success" ? "80" : roleName == "Warning" ? "60" : "40")),
            reopened.Slides[0].Tables.Single().GetCell(0, 0).FillColor);
    }

    [Theory]
    [InlineData("master")]
    [InlineData("layout")]
    [InlineData("slide")]
    [InlineData("slide-master")]
    [InlineData("layout-master")]
    public void ContrastUsesTheSavedNativeColorMapAndTextInheritance(string scope) {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        document.SetThemeColor(PowerPointThemeColor.Dark1, "000000");
        document.SetThemeColor(PowerPointThemeColor.Accent1, "000000");
        document.SetThemeColor(PowerPointThemeColor.Accent2, "FFFFFF");
        foreach (A.SchemeColorValues slot in new[] { A.SchemeColorValues.Accent1, A.SchemeColorValues.Text1, A.SchemeColorValues.Background1 }) {
            foreach (string source in new[] { "run", "paragraph", "font-reference" }) {
                var box = slide.AddTextBox(slot + " " + source);
                box.FillColor = OfficeColor.White;
                box.FontSize = OfficeLength.Points(12);
                var paragraph = box.Paragraphs[0];
                var fill = new A.SolidFill(new A.SchemeColor { Val = slot });
                if (source == "run") paragraph.Runs[0].Run.RunProperties!.AddChild(fill, true);
                else if (source == "paragraph") {
                    paragraph.Paragraph.ParagraphProperties ??= new A.ParagraphProperties();
                    paragraph.Paragraph.ParagraphProperties.AddChild(new A.DefaultRunProperties(fill), true);
                } else {
                    ((P.Shape)box.Element).ShapeStyle = new P.ShapeStyle(
                        new A.LineReference { Index = 0U }, new A.FillReference { Index = 0U },
                        new A.EffectReference { Index = 0U }, new A.FontReference(new A.SchemeColor { Val = slot }) { Index = A.FontCollectionIndexValues.Minor });
                }
            }
        }
        var fieldBox = slide.AddTextBox(string.Empty);
        fieldBox.FillColor = OfficeColor.White;
        fieldBox.Paragraphs[0].AddField("Field", "slidenum", null, run => {
            run.FontSize = OfficeLength.Points(12);
            run.RunProperties!.AddChild(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Text1 }), true);
        });
        var mappedBackground = slide.AddTextBox("Mapped background");
        mappedBackground.SetColor(OfficeColor.White).SetFontSize(OfficeLength.Points(12));
        ((P.Shape)mappedBackground.Element).ShapeProperties!.AddChild(
            new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Background1 }), true);
        SetMap(slide, scope, A.ColorSchemeIndexValues.Accent2);
        using (var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()))) {
            var report = saved.InspectAccessibility(PowerPointAccessibilityOptions.ForProfile(PowerPointAccessibilityPolicyProfile.Strict));
            var findings = report.Findings.Where(finding => finding.Code == "accessibility-low-contrast").ToArray();
            Assert.Equal(11, findings.Length);
            Assert.All(findings, finding => Assert.Equal(1D, finding.MeasuredValue));
        }
        SetMap(slide, scope, A.ColorSchemeIndexValues.Dark1);
        using var control = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.DoesNotContain(control.InspectAccessibility(PowerPointAccessibilityOptions.ForProfile(PowerPointAccessibilityPolicyProfile.Strict)).Findings,
            finding => finding.Code == "accessibility-low-contrast");
    }

    [Fact]
    public void TypedDesignPaletteKeepsPresetsAndResolvesRolesInThePresentation() {
        var design = PowerPointDesignTheme.FromBrand("#7A3E9D");
        Assert.Equal(OfficeColor.ParseHex("7A3E9D"), design.AccentColor);
        Assert.Equal(OfficeColor.ParseHex("FFFFFF"), design.BackgroundColor);
        design.PanelColor = OfficeColor.ParseHex("12345680");
        design.AccentColor = OfficeColor.Accent4;
        var clone = design.Clone();
        Assert.Equal(OfficeColor.Accent4, clone.AccentColor);
        Assert.Equal(OfficeColor.ParseHex("12345680"), clone.PanelColor);
        var theme = OfficeTheme.Default.WithPalette(OfficeTheme.Default.Palette.With(accent4: OfficeColor.ParseHex("003366")));
        using var document = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = theme });
        document.ApplyDesignerTheme(clone);
        document.AddSlide().AddTextBox("Panel").FillColor = clone.PanelColor;
        using var native = PresentationDocument.Open(new MemoryStream(document.ToBytes()), false);
        Assert.Equal("003366", native.PresentationPart!.SlideMasterParts.First().ThemePart!.Theme.ThemeElements!.ColorScheme!
            .Accent1Color!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
        AssertPaint(native.PresentationPart.SlideParts.Single().Slide.Descendants<P.Shape>().Single().ShapeProperties!, "123456", 50196);
    }

    private static void SetMap(PowerPointSlide slide, string scope, A.ColorSchemeIndexValues target) {
        var layout = slide.SlidePart.SlideLayoutPart!;
        var master = layout.SlideMasterPart!.SlideMaster;
        P.ColorMap map = master.ColorMap!;
        map.Accent1 = map.Text1 = map.Background1 = scope is "master" or "slide-master" or "layout-master" ? target : A.ColorSchemeIndexValues.Dark1;
        var nativeOverride = new A.OverrideColorMapping();
        nativeOverride.SetAttributes(map.GetAttributes());
        nativeOverride.Accent1 = nativeOverride.Text1 = nativeOverride.Background1 = target;
        layout.SlideLayout.ColorMapOverride = scope == "layout" ? new P.ColorMapOverride(nativeOverride)
            : scope == "slide-master" ? new P.ColorMapOverride(BlackMap(map))
            : scope == "layout-master" ? new P.ColorMapOverride(new A.MasterColorMapping()) : null;
        slide.SlidePart.Slide.ColorMapOverride = scope == "slide" ? new P.ColorMapOverride(nativeOverride)
            : scope == "slide-master" ? new P.ColorMapOverride(new A.MasterColorMapping()) : null;
    }

    private static A.OverrideColorMapping BlackMap(P.ColorMap source) {
        var map = new A.OverrideColorMapping();
        map.SetAttributes(source.GetAttributes());
        map.Accent1 = map.Text1 = map.Background1 = A.ColorSchemeIndexValues.Dark1;
        return map;
    }

    private static void AssertPaint(OpenXmlElement container, string rgb, int alpha) =>
        AssertColor(container.GetFirstChild<A.SolidFill>()!.FirstChild!, rgb, alpha);

    private static void AssertColor(OpenXmlElement color, string rgb, int alpha) {
        var native = Assert.IsType<A.RgbColorModelHex>(color);
        Assert.Equal(rgb, native.Val!.Value);
        Assert.Equal(alpha, Assert.Single(native.Elements<A.Alpha>()).Val!.Value);
    }
}
