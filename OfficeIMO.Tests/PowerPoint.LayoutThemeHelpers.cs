using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointLayoutThemeHelpersTests {
        [Fact]
        public void CanSelectLayoutsByTypeAndName() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                int titleOnlyIndex = presentation.GetLayoutIndex(SlideLayoutValues.TitleOnly);

                PowerPointSlide slide = presentation.AddSlide(SlideLayoutValues.TitleOnly);
                Assert.Equal(titleOnlyIndex, slide.LayoutIndex);

                int titleSlideIndex = presentation.GetLayoutIndex("Title Slide");
                slide.SetLayout("Title Slide");
                Assert.Equal(titleSlideIndex, slide.LayoutIndex);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanSetThemeColorsAndFonts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.SetThemeColor(PowerPointThemeColor.Accent1, "FF0000");
                    presentation.SetThemeLatinFonts("Aptos", "Calibri");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlideMasterPart master = document.PresentationPart!.SlideMasterParts.First();
                    A.ColorScheme? scheme = master.ThemePart?.Theme?.ThemeElements?.ColorScheme;
                    Assert.NotNull(scheme);

                    string? accent1 = scheme!.GetFirstChild<A.Accent1Color>()
                        ?.GetFirstChild<A.RgbColorModelHex>()
                        ?.Val?.Value;
                    Assert.Equal("FF0000", accent1);

                    A.FontScheme? fontScheme = master.ThemePart?.Theme?.ThemeElements?.FontScheme;
                    Assert.NotNull(fontScheme);
                    Assert.Equal("Aptos", fontScheme!.MajorFont?.LatinFont?.Typeface);
                    Assert.Equal("Calibri", fontScheme.MinorFont?.LatinFont?.Typeface);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanSetThemeFontsAcrossScripts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.SetThemeFonts(new PowerPointThemeFontSet(
                        majorLatin: "Aptos",
                        minorLatin: "Calibri",
                        majorEastAsian: "MS Mincho",
                        minorEastAsian: "Yu Gothic",
                        majorComplexScript: "Arial",
                        minorComplexScript: "Tahoma"));
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlideMasterPart master = document.PresentationPart!.SlideMasterParts.First();
                    A.FontScheme? fontScheme = master.ThemePart?.Theme?.ThemeElements?.FontScheme;
                    Assert.NotNull(fontScheme);
                    Assert.Equal("Aptos", fontScheme!.MajorFont?.LatinFont?.Typeface);
                    Assert.Equal("Calibri", fontScheme.MinorFont?.LatinFont?.Typeface);
                    Assert.Equal("MS Mincho", fontScheme.MajorFont?.EastAsianFont?.Typeface);
                    Assert.Equal("Yu Gothic", fontScheme.MinorFont?.EastAsianFont?.Typeface);
                    Assert.Equal("Arial", fontScheme.MajorFont?.ComplexScriptFont?.Typeface);
                    Assert.Equal("Tahoma", fontScheme.MinorFont?.ComplexScriptFont?.Typeface);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
