using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests;

public sealed class PowerPointCreationThemeTests {
    private static OfficeTheme CustomTheme => OfficeTheme.Default.WithPalette(
        OfficeTheme.Default.Palette.With(accent1: OfficeColor.ParseHex("123456"))).WithFonts(
        OfficeTheme.Default.Fonts.With(heading: OfficeTheme.Default.Fonts.Heading.With(family: "Georgia")));

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void CreationOverloadsWriteThemeToNativePartsAndReload(int destination) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
        using var stream = new MemoryStream();
        try {
            var options = new PowerPointCreateOptions { Theme = CustomTheme };
            using (var document = destination switch {
                1 => PowerPointDocument.Create(path, options),
                2 => PowerPointDocument.Create(stream, options),
                _ => PowerPointDocument.Create(options)
            }) {
                Assert.Equal("123456", document.OpenXmlDocument.PresentationPart!.SlideMasterParts.First()
                    .ThemePart!.Theme.ThemeElements!.ColorScheme!.Accent1Color!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
                Assert.Equal("Georgia", document.OpenXmlDocument.PresentationPart.SlideMasterParts.First()
                    .ThemePart!.Theme.ThemeElements!.FontScheme!.MajorFont!.LatinFont!.Typeface!.Value);
                document.AddSlide().AddTitle("Native theme");
                if (destination == 0) document.Save(stream); else document.Save();
            }
            if (destination == 1) stream.Write(File.ReadAllBytes(path), 0, checked((int)new FileInfo(path).Length));
            stream.Position = 0;
            using var loaded = PowerPointDocument.Load(stream);
            Assert.Equal("123456", loaded.OpenXmlDocument.PresentationPart!.SlideMasterParts.First()
                .ThemePart!.Theme.ThemeElements!.ColorScheme!.Accent1Color!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void NullThemeUsesCanonicalDefaultAndNeutralImportDoesNotOwnTheme() {
        using var authored = PowerPointDocument.Create();
        using var imported = PowerPointDocument.CreateForImport();
        Assert.True(authored.OwnsLayoutTheme);
        Assert.Same(OfficeTheme.Default, authored.BaseTheme);
        Assert.False(imported.OwnsLayoutTheme);
        Assert.NotEqual(authored.OpenXmlDocument.PresentationPart!.Presentation.DefaultTextStyle!.OuterXml,
            imported.OpenXmlDocument.PresentationPart!.Presentation.DefaultTextStyle!.OuterXml);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void RepeatedCreationKeepsMutablePackagesAndThemesIndependent(int creation) {
        PowerPointDocument CreatePresentation() => creation switch {
            1 => PowerPointDocument.Create(new PowerPointCreateOptions { Theme = CustomTheme }),
            2 => PowerPointDocument.CreateForImport(),
            _ => PowerPointDocument.Create()
        };

        using var first = CreatePresentation();
        string? originalAccent = first.GetThemeColor(PowerPointThemeColor.Accent1);
        var firstMaster = first.OpenXmlDocument.PresentationPart!.SlideMasterParts.First().SlideMaster;
        string originalMaster = firstMaster.OuterXml;
        DateTime oldCreated = new(2000, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        firstMaster.CommonSlideData!.Name = "First master";
        first.SetThemeColor(PowerPointThemeColor.Accent1, "FF00FF");
        first.Properties.Title = "First package";
        first.Properties.Created = oldCreated;
        first.ApplicationProperties.Company = "First company";
        first.AddSlide().AddTitle("First slide");

        using var second = CreatePresentation();
        Assert.Empty(second.Slides);
        Assert.Equal(originalAccent, second.GetThemeColor(PowerPointThemeColor.Accent1));
        Assert.Equal(originalMaster, second.OpenXmlDocument.PresentationPart!.SlideMasterParts.First().SlideMaster.OuterXml);
        Assert.True(string.IsNullOrEmpty(second.Properties.Title));
        Assert.NotEqual(oldCreated, second.Properties.Created);
        Assert.Equal(string.Empty, second.ApplicationProperties.Company);
        second.AddSlide().AddTitle("Second slide");

        using var firstStream = new MemoryStream();
        first.Save(firstStream);
        first.Dispose();
        using var secondStream = new MemoryStream();
        second.Save(secondStream);
        using var firstSaved = PowerPointDocument.Load(firstStream);
        using var secondSaved = PowerPointDocument.Load(secondStream);
        Assert.Equal("First slide", Assert.Single(firstSaved.Slides).TextBoxes.First().Text);
        Assert.Equal("FF00FF", firstSaved.GetThemeColor(PowerPointThemeColor.Accent1));
        Assert.Equal(oldCreated, firstSaved.Properties.Created?.ToUniversalTime());
        Assert.Equal("Second slide", Assert.Single(secondSaved.Slides).TextBoxes.First().Text);
        Assert.Equal(originalAccent, secondSaved.GetThemeColor(PowerPointThemeColor.Accent1));
        Assert.True(string.IsNullOrEmpty(secondSaved.Properties.Title));
        Assert.Equal(string.Empty, secondSaved.ApplicationProperties.Company);

        using var fresh = CreatePresentation();
        Assert.Empty(fresh.Slides);
        Assert.Equal(originalAccent, fresh.GetThemeColor(PowerPointThemeColor.Accent1));
        Assert.True(string.IsNullOrEmpty(fresh.Properties.Title));
        Assert.Empty(firstSaved.ValidateDocument());
        Assert.Empty(secondSaved.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TemplateStyleSurvivesUnlessThemeIsExplicitlyChosen(bool chooseTheme) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
        try {
            using (var template = PowerPointDocument.CreateForImport()) {
                var native = template.OpenXmlDocument.PresentationPart!.SlideMasterParts.First();
                native.ThemePart!.Theme.ThemeElements!.ColorScheme!.Accent1Color!.RemoveAllChildren();
                native.ThemePart.Theme.ThemeElements.ColorScheme.Accent1Color.Append(new A.RgbColorModelHex { Val = "AB12CD" });
                template.Save(path);
            }
            byte[] original = File.ReadAllBytes(path);
            using var created = PowerPointDocument.Create(new PowerPointCreateOptions {
                Template = path, Theme = chooseTheme ? CustomTheme : null
            });
            Assert.Equal(chooseTheme ? "123456" : "AB12CD", created.OpenXmlDocument.PresentationPart!.SlideMasterParts.First()
                .ThemePart!.Theme.ThemeElements!.ColorScheme!.Accent1Color!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(chooseTheme, created.OwnsLayoutTheme);
            Assert.Equal(original, File.ReadAllBytes(path));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ReadingLoadedThemeTokensDoesNotRewriteNativeStyleParts() {
        using var original = PowerPointDocument.CreateForImport();
        using var stream = new MemoryStream();
        original.Save(stream);
        stream.Position = 0;
        using var loaded = PowerPointDocument.Load(stream);
        string before = loaded.OpenXmlDocument.PresentationPart!.SlideMasterParts.First().SlideMaster.OuterXml;
        string theme = loaded.OpenXmlDocument.PresentationPart.SlideMasterParts.First().ThemePart!.Theme.OuterXml;
        _ = loaded.LayoutTheme;
        loaded.AddSlide().AddChart(OfficeChartType.Column, Data());
        Assert.False(loaded.OwnsLayoutTheme);
        Assert.Equal(before, loaded.OpenXmlDocument.PresentationPart.SlideMasterParts.First().SlideMaster.OuterXml);
        Assert.Equal(theme, loaded.OpenXmlDocument.PresentationPart.SlideMasterParts.First().ThemePart!.Theme.OuterXml);
    }

    [Fact]
    public void NeutralStreamImportRetainsExplicitSaveDestinationAndCallerStreamLifetime() {
        using var stream = new MemoryStream();
        using (var imported = PowerPointDocument.CreateForImport(stream)) {
            Assert.False(imported.OwnsLayoutTheme);
            imported.AddSlide();
            Assert.Equal(0L, stream.Length);
            imported.Save();
            Assert.True(stream.Length > 0L);
        }
        Assert.True(stream.CanWrite);
    }

    private static OfficeChartData Data() => new(new[] { "Q1", "Q2" }, new[] { new OfficeSeries("Revenue", new[] { 12D, 17D }) });

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ThemeEditsReachNewSavedChartsAfterTheAuthoringCacheIsUsed(bool reopen) {
        OfficeTheme theme = CustomTheme.WithChartStyle(CustomTheme.ChartStyle.With(axisFontSize: OfficeLength.Points(17)));
        using var authored = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = theme });
        authored.AddSlide();
        using var loaded = reopen ? PowerPointDocument.Load(new MemoryStream(authored.ToBytes())) : null;
        PowerPointDocument document = loaded ?? authored;
        var bounds = new PowerPointChartOptions { Left = OfficeLength.Points(40), Top = OfficeLength.Points(80),
            Width = OfficeLength.Points(160), Height = OfficeLength.Points(120) };
        document.Slides[0].AddChart(OfficeChartType.Column, Data(), bounds);
        document.SetThemeColor(PowerPointThemeColor.Accent1, "FF00FF");
        document.SetThemeColor(PowerPointThemeColor.Light2, "113355");
        document.SetThemeLatinFonts("Cambria", "Calibri");
        document.Slides[0].AddChart(OfficeChartType.Column, Data(), bounds);
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        var charts = saved.OpenXmlDocument.PresentationPart!.SlideParts.Single().ChartParts.ToArray();
        Assert.Equal(2, charts.Length);
        Assert.Contains(charts, part => part.ChartSpace.OuterXml.Contains("FF00FF"));
        Assert.Equal("Cambria", saved.GetThemeFonts().MajorLatin);
        Assert.Equal("Calibri", saved.GetThemeFonts().MinorLatin);
        Assert.Equal("113355", document.GetTheme().Surfaces.Surface.ToRgbHex(), ignoreCase: true);
        // A palette/font edit must not reset an authored deck's unrelated chart tokens.
        if (!reopen) Assert.Equal(17, document.GetTheme().ChartStyle.AxisFontSize.ToPoints());
        Assert.Empty(saved.ValidateDocument());
    }

    [Theory]
    [InlineData("series")]
    [InlineData("point")]
    [InlineData("marker")]
    [InlineData("series-style")]
    [InlineData("point-hatch")]
    [InlineData("series-muted")]
    public void SemanticChartPaintIsResolvedAgainstTheDeckThemeBeforeNativeSerialization(string paint) {
        using var document = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = CustomTheme });
        if (paint == "series-muted") document.SetThemeColor(PowerPointThemeColor.Dark2, "123456");
        var series = new OfficeSeries("Revenue", new[] { 12D, 17D }, null,
            paint == "series" ? OfficeColor.Accent1 : paint == "series-muted" ? OfficeColor.TextMuted : null,
            paint == "point" ? new OfficeColor?[] { OfficeColor.Accent1, null } : null,
            showMarkers: true, markerOutlineColor: paint == "marker" ? OfficeColor.Accent1 : null);
        if (paint == "series-style") series = series.WithSeriesStyle(new OfficeChartPointStyle(fillColor: OfficeColor.Accent1));
        if (paint == "point-hatch") series = series.WithPointStyles(new OfficeChartPointStyle?[] {
            new(hatch: OfficeChartHatchPattern.Cross, hatchColor: OfficeColor.Accent1), null });
        document.AddSlide().AddChart(paint == "marker" ? OfficeChartType.Line : OfficeChartType.Column,
            new OfficeChartData(new[] { "Q1", "Q2" }, new[] { series }));
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        var native = saved.OpenXmlDocument.PresentationPart!.SlideParts.Single().ChartParts.Single().ChartSpace;
        Assert.Contains(native.Descendants<A.RgbColorModelHex>(), color => color.Val?.Value == "123456");
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void ChartCreationAndDataUpdatesUseTheSlidesOwnMasterTheme() {
        using var document = PowerPointDocument.Create();
        using var source = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = CustomTheme });
        var part = document.OpenXmlDocument.PresentationPart!;
        var master = part.AddPart(source.OpenXmlDocument.PresentationPart!.SlideMasterParts.First());
        part.Presentation.SlideMasterIdList!.Append(new DocumentFormat.OpenXml.Presentation.SlideMasterId {
            Id = 2147483649U, RelationshipId = part.GetIdOfPart(master) });
        var data = new OfficeChartData(new[] { "Q1", "Q2" }, new[] {
            new OfficeSeries("Revenue", new[] { 12D, 17D }, null, OfficeColor.Accent1) });
        document.ApplyTheme(CustomTheme.WithName("Second master"), 1);
        document.SetThemeColor(PowerPointThemeColor.Accent1, "654321", 1);
        var chart = document.AddSlide(masterIndex: 1).AddChart(OfficeChartType.Column, data);
        Assert.True(chart.TryGetOfficeSnapshot(out var initial));
        Assert.Equal("654321", initial.Data.Series[0].Color!.Value.ToRgbHex(), ignoreCase: true);
        document.SetThemeColor(PowerPointThemeColor.Accent1, "123456", 1);
        chart.UpdateData(data);
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.True(saved.Slides.Single().Charts.Single().TryGetOfficeSnapshot(out var updated));
        Assert.Equal("123456", updated.Data.Series[0].Color!.Value.ToRgbHex(), ignoreCase: true);
        Assert.Equal("Second master", saved.GetTheme(1).Name);
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void ApplyThemeOnLoadedDeckKeepsMasterAndLayoutContent() {
        using var stream = new MemoryStream();
        using (var author = PowerPointDocument.Create()) {
            author.AddSlide();
            author.Save(stream);
        }

        stream.Position = 0;
        using (var package = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(stream, true)) {
            var master = package.PresentationPart!.SlideMasterParts.First();
            master.SlideMaster.CommonSlideData!.ShapeTree!.Append(BrandShape(99U, "MasterBrand"));
            master.SlideMaster.CommonSlideData.Background!.BackgroundProperties!.RemoveAllChildren<A.SolidFill>();
            master.SlideMaster.CommonSlideData.Background.BackgroundProperties.PrependChild(new A.SolidFill(new A.RgbColorModelHex { Val = "ABCDEF" }));
            master.SlideMaster.Save();
            var layout = master.SlideLayoutParts.First(part => part.SlideLayout.CommonSlideData!.Name == "Title and Content");
            layout.SlideLayout.CommonSlideData!.ShapeTree!.Append(BrandShape(98U, "LayoutBrand"));
            layout.SlideLayout.Save();
        }

        stream.Position = 0;
        using var loaded = PowerPointDocument.Load(stream);
        string layoutsBefore = string.Concat(loaded.OpenXmlDocument.PresentationPart!.SlideMasterParts.First()
            .SlideLayoutParts.Select(part => part.SlideLayout.OuterXml));
        loaded.ApplyTheme(CustomTheme);
        loaded.ApplyTheme(CustomTheme.WithName("Second pass"));

        Assert.False(loaded.OwnsLayoutTheme);
        using var saved = PowerPointDocument.Load(new MemoryStream(loaded.ToBytes()));
        var savedMaster = saved.OpenXmlDocument.PresentationPart!.SlideMasterParts.First();
        Assert.Contains(savedMaster.SlideMaster.Descendants<DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties>(),
            shape => shape.Name == "MasterBrand");
        Assert.Equal("ABCDEF", savedMaster.SlideMaster.CommonSlideData!.Background!.BackgroundProperties!
            .GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
        Assert.Equal(layoutsBefore, string.Concat(savedMaster.SlideLayoutParts.Select(part => part.SlideLayout.OuterXml)));
        Assert.Equal("123456", savedMaster.ThemePart!.Theme.ThemeElements!.ColorScheme!.Accent1Color!
            .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
        Assert.Equal("Second pass", saved.GetTheme().Name);
        Assert.Contains("lnSpc", savedMaster.SlideMaster.TextStyles!.TitleStyle!.OuterXml);
        Assert.Empty(saved.ValidateDocument());
    }

    private static DocumentFormat.OpenXml.Presentation.Shape BrandShape(uint id, string name) => new(
        new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
            new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = id, Name = name },
            new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(),
            new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()),
        new DocumentFormat.OpenXml.Presentation.ShapeProperties());
}
