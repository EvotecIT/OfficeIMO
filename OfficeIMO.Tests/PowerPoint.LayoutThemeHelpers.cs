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
        public void SetThemeLatinFontsPreservesExistingScriptFontAttributes() {
            string filePath = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(Path.GetRandomFileName(), ".pptx"));
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.SetThemeFonts(new PowerPointThemeFontSet(
                        majorLatin: "Major Latin",
                        minorLatin: "Minor Latin",
                        majorEastAsian: "Major East Asian",
                        minorEastAsian: "Minor East Asian",
                        majorComplexScript: "Major Complex",
                        minorComplexScript: "Minor Complex"));
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlideMasterPart master = document.PresentationPart!.SlideMasterParts.First();
                    A.FontScheme fontScheme = master.ThemePart!.Theme!.ThemeElements!.FontScheme!;
                    A.EastAsianFont eastAsian = fontScheme.MajorFont!.EastAsianFont!;
                    eastAsian.Panose = "020B0604020202020204";
                    eastAsian.PitchFamily = 34;
                    eastAsian.CharacterSet = 0;

                    A.ComplexScriptFont complexScript = fontScheme.MajorFont.ComplexScriptFont!;
                    complexScript.Panose = "020B0604020202020205";
                    complexScript.PitchFamily = 18;
                    complexScript.CharacterSet = 1;
                    master.ThemePart.Theme.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    presentation.SetThemeLatinFonts("Updated Major Latin", "Updated Minor Latin");
                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlideMasterPart master = document.PresentationPart!.SlideMasterParts.First();
                    A.FontScheme fontScheme = master.ThemePart!.Theme!.ThemeElements!.FontScheme!;

                    Assert.Equal("Updated Major Latin", fontScheme.MajorFont!.LatinFont!.Typeface);
                    A.EastAsianFont eastAsian = fontScheme.MajorFont.EastAsianFont!;
                    Assert.Equal("Major East Asian", eastAsian.Typeface);
                    Assert.Equal("020B0604020202020204", eastAsian.Panose);
                    Assert.Equal(34, eastAsian.PitchFamily);
                    Assert.Equal(0, eastAsian.CharacterSet);

                    A.ComplexScriptFont complexScript = fontScheme.MajorFont.ComplexScriptFont!;
                    Assert.Equal("Major Complex", complexScript.Typeface);
                    Assert.Equal("020B0604020202020205", complexScript.Panose);
                    Assert.Equal(18, complexScript.PitchFamily);
                    Assert.Equal(1, complexScript.CharacterSet);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DefaultThemeColors_IncludeSystemColorBackedEntries() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);

                IReadOnlyDictionary<PowerPointThemeColor, string> colors = presentation.GetThemeColors();

                Assert.Equal("000000", presentation.GetThemeColor(PowerPointThemeColor.Dark1));
                Assert.Equal("FFFFFF", presentation.GetThemeColor(PowerPointThemeColor.Light1));
                Assert.Equal("000000", colors[PowerPointThemeColor.Dark1]);
                Assert.Equal("FFFFFF", colors[PowerPointThemeColor.Light1]);
                Assert.Equal("156082", colors[PowerPointThemeColor.Accent1]);
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
                    Assert.Empty(presentation.ValidateDocument());
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
                    Assert.IsType<A.LatinFont>(fontScheme.MajorFont!.ChildElements[0]);
                    Assert.IsType<A.EastAsianFont>(fontScheme.MajorFont.ChildElements[1]);
                    Assert.IsType<A.ComplexScriptFont>(fontScheme.MajorFont.ChildElements[2]);
                    Assert.IsType<A.LatinFont>(fontScheme.MinorFont!.ChildElements[0]);
                    Assert.IsType<A.EastAsianFont>(fontScheme.MinorFont.ChildElements[1]);
                    Assert.IsType<A.ComplexScriptFont>(fontScheme.MinorFont.ChildElements[2]);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
