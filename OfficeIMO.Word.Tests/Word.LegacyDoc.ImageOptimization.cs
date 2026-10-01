using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData("eastAsia")]
        [InlineData("complexScript")]
        [InlineData("highAnsi")]
        public void LegacyDoc_DefaultStyleRetainsOneScriptSpecificFont(string selector) {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Default style font");
            Style normal = GetNormalStyle(document);
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.RemoveAllChildren<RunFonts>();
            var fonts = new RunFonts();
            if (selector == "eastAsia") fonts.EastAsia = "Courier New";
            else if (selector == "complexScript") fonts.ComplexScript = "Courier New";
            else fonts.HighAnsi = "Courier New";
            normal.StyleRunProperties = new StyleRunProperties(fonts);
            using var output = new MemoryStream(document.ToBytes(WordFileFormat.Doc));
            using WordDocument reopened = WordDocument.Load(output);
            Assert.Equal("Courier New", GetNormalStyle(reopened).StyleRunProperties?.GetFirstChild<RunFonts>()?.Ascii?.Value);
        }

        [Fact]
        public void LegacyDoc_DefaultStyleRejectsDistinctExplicitFontSelectors() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Mixed default fonts");
            GetNormalStyle(document).StyleRunProperties = new StyleRunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Courier New" });
            Assert.Throws<NotSupportedException>(() => document.ToBytes(WordFileFormat.Doc));
        }

        [Theory]
        [InlineData("Courier New", false)]
        [InlineData("Arial", true)]
        public void LegacyDoc_DefaultStyleValidatesInheritedAndOverriddenFontSelectors(string overrideFont, bool supported) {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Inherited Latin font");
            var defaults = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!;
            defaults.RemoveAllChildren<RunFonts>();
            defaults.PrependChild(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" });
            GetNormalStyle(document).StyleRunProperties = new StyleRunProperties(new RunFonts { EastAsia = overrideFont });
            if (!supported) Assert.Throws<NotSupportedException>(() => document.ToBytes(WordFileFormat.Doc));
            else {
                using var output = new MemoryStream(document.ToBytes(WordFileFormat.Doc));
                using WordDocument reopened = WordDocument.Load(output);
                Assert.Equal("Arial", GetNormalStyle(reopened).StyleRunProperties?.GetFirstChild<RunFonts>()?.Ascii?.Value);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void LegacyDoc_DefaultStyleResolvesMajorAsciiAndHighAnsiThemeFonts(bool ascii) {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Theme default font");
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.RemoveAllChildren<RunFonts>();
            var main = document._wordprocessingDocument.MainDocumentPart!;
            var theme = main.ThemePart ?? main.AddNewPart<DocumentFormat.OpenXml.Packaging.ThemePart>();
            theme.Theme = new A.Theme(new A.ThemeElements(new A.FontScheme(
                new A.MajorFont(new A.LatinFont { Typeface = "Courier New" }),
                new A.MinorFont(new A.LatinFont { Typeface = "Arial" })) { Name = "Test fonts" }));
            var fonts = new RunFonts();
            if (ascii) fonts.AsciiTheme = ThemeFontValues.MajorAscii;
            else fonts.HighAnsiTheme = ThemeFontValues.MajorHighAnsi;
            GetNormalStyle(document).StyleRunProperties = new StyleRunProperties(fonts);
            using var output = new MemoryStream(document.ToBytes(WordFileFormat.Doc));
            using WordDocument reopened = WordDocument.Load(output);
            Assert.Equal("Courier New", GetNormalStyle(reopened).StyleRunProperties?.GetFirstChild<RunFonts>()?.Ascii?.Value);
        }

        private static Style GetNormalStyle(WordDocument document) => Assert.Single(
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>(),
            style => style.StyleId?.Value == "Normal");

        [Fact]
        public void LegacyDoc_DefaultStyleRejectsDistinctResolvedThemeFonts() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Mixed theme defaults");
            var main = document._wordprocessingDocument.MainDocumentPart!;
            var theme = main.ThemePart ?? main.AddNewPart<DocumentFormat.OpenXml.Packaging.ThemePart>();
            theme.Theme = new A.Theme(new A.ThemeElements(new A.FontScheme(
                new A.MajorFont(new A.LatinFont { Typeface = "Arial" }),
                new A.MinorFont(new A.LatinFont { Typeface = "Courier New" })) { Name = "Distinct fonts" }));
            GetNormalStyle(document).StyleRunProperties = new StyleRunProperties(new RunFonts {
                AsciiTheme = ThemeFontValues.MajorAscii, HighAnsiTheme = ThemeFontValues.MinorHighAnsi
            });
            Assert.Throws<NotSupportedException>(() => document.ToBytes(WordFileFormat.Doc));
        }

        [Fact]
        public void LegacyDoc_WordProducedInlineCropAndDefaultFontSurviveImageOptimization() {
            using WordDocument document = WordDocument.Load(GetFixtureDoc(Path.Combine("LegacyDocCorpus", "ComCroppedInlinePicture.doc")));
            Assert.Empty(document.LegacyDocUnsupportedFeatures);
            var picture = Assert.Single(document.Images);
            int? crop = picture.CropLeft;
            Assert.True(crop > 0);
            AssertNormalFont(document);
            var report = document.OptimizeImages(new WordImageOptimizationOptions {
                Mode = OfficeImageOptimizationMode.DownsampleAndRecompress, TargetDpi = 144, JpegQuality = 65
            });
            Assert.True(report.BytesSaved > 0);
            using var output = new MemoryStream(document.ToBytes(WordFileFormat.Doc));
            using WordDocument reopened = WordDocument.Load(output);
            AssertNormalFont(reopened);
            Assert.Contains("Independent Word producer - cropped image optimization.", reopened.Paragraphs.Select(p => p.Text));
            Assert.Equal(crop, Assert.Single(reopened.Images).CropLeft);
            Assert.Empty(reopened.ValidateDocument());
        }

        private static void AssertNormalFont(WordDocument document) {
            var normal = Assert.Single(document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>(),
                style => style.StyleId?.Value == "Normal");
            Assert.Equal("Aptos", normal.StyleRunProperties?.GetFirstChild<RunFonts>()?.Ascii?.Value);
            Assert.Equal("24", normal.StyleRunProperties?.GetFirstChild<FontSize>()?.Val?.Value);
        }
    }
}
