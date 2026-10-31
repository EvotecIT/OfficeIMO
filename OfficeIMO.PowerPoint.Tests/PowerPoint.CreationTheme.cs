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
}
