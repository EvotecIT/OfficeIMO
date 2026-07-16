using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptRoundTripThemeTests {
        private static string AccessibilityFixture => Path.Combine(
            AppContext.BaseDirectory, "Documents", "LegacyPptCorpus",
            "AccessibilityPowerPoint.ppt");

        [Fact]
        public void BinaryImport_DecodesAndProjectsCompleteRoundTripTheme() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(
                AccessibilityFixture);
            int masterIndex = legacy.Masters.ToList().FindIndex(master =>
                master.RoundTripTheme != null);
            Assert.True(masterIndex >= 0);
            LegacyPptRoundTripTheme theme = Assert.IsType<
                LegacyPptRoundTripTheme>(legacy.Masters[masterIndex]
                .RoundTripTheme);

            Assert.False(theme.IsOverride);
            Assert.Equal("Office Theme", theme.Name);
            Assert.Equal("Office", theme.ColorSchemeName);
            Assert.Equal("4BACC6",
                theme.Colors[PowerPointThemeColor.Accent5]);
            Assert.Equal("F79646",
                theme.Colors[PowerPointThemeColor.Accent6]);
            Assert.Equal("0000FF",
                theme.Colors[PowerPointThemeColor.Hyperlink]);
            Assert.Equal("800080",
                theme.Colors[PowerPointThemeColor.FollowedHyperlink]);
            Assert.Equal("Calibri", theme.MajorLatinFont);
            Assert.Equal("Calibri", theme.MinorLatinFont);
            Assert.Contains("<a:fmtScheme", theme.ThemeXml);
            Assert.Contains("<a:clrMap", theme.ColorMappingXml);
            Assert.True(legacy.CreateImportReport().RoundTripThemeCount > 0);

            using PowerPointPresentation projected = PowerPointPresentation.Load(
                AccessibilityFixture);
            Assert.Equal("4BACC6", projected.GetThemeColor(
                PowerPointThemeColor.Accent5, masterIndex));
            Assert.Equal("F79646", projected.GetThemeColor(
                PowerPointThemeColor.Accent6, masterIndex));
            Assert.Equal("0000FF", projected.GetThemeColor(
                PowerPointThemeColor.Hyperlink, masterIndex));
            Assert.Equal("800080", projected.GetThemeColor(
                PowerPointThemeColor.FollowedHyperlink, masterIndex));
            Assert.Equal("Calibri",
                projected.GetThemeFonts(masterIndex).MajorLatin);
            Assert.NotNull(projected.OpenXmlDocument.PresentationPart!
                .SlideMasterParts.ElementAt(masterIndex).ThemePart?.Theme?
                .ThemeElements?.FormatScheme);
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_RoundTripsCompleteDrawingMlThemePackage() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.SetThemeNameForAllMasters("OfficeIMO Binary Theme");
            presentation.SetThemeColorsForAllMasters(new Dictionary<
                PowerPointThemeColor, string> {
                    [PowerPointThemeColor.Accent5] = "102030",
                    [PowerPointThemeColor.Accent6] = "405060",
                    [PowerPointThemeColor.Hyperlink] = "708090",
                    [PowerPointThemeColor.FollowedHyperlink] = "A0B0C0"
                });
            presentation.SetThemeLatinFontsForAllMasters(
                "Aptos Display", "Aptos");
            presentation.AddSlide(P.SlideLayoutValues.Blank);

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptRoundTripTheme written = Assert.IsType<
                LegacyPptRoundTripTheme>(Assert.Single(
                    LegacyPptPresentation.Load(bytes).Masters).RoundTripTheme);

            Assert.Equal("OfficeIMO Binary Theme", written.Name);
            Assert.Equal("102030",
                written.Colors[PowerPointThemeColor.Accent5]);
            Assert.Equal("405060",
                written.Colors[PowerPointThemeColor.Accent6]);
            Assert.Equal("708090",
                written.Colors[PowerPointThemeColor.Hyperlink]);
            Assert.Equal("A0B0C0",
                written.Colors[PowerPointThemeColor.FollowedHyperlink]);
            Assert.Equal("Aptos Display", written.MajorLatinFont);
            Assert.Equal("Aptos", written.MinorLatinFont);
            Assert.Contains("<a:fmtScheme", written.ThemeXml);
            Assert.Contains("<a:clrMap", written.ColorMappingXml);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(stream);
            Assert.Equal("102030", reopened.GetThemeColor(
                PowerPointThemeColor.Accent5));
            Assert.Equal("405060", reopened.GetThemeColor(
                PowerPointThemeColor.Accent6));
            Assert.Equal("708090", reopened.GetThemeColor(
                PowerPointThemeColor.Hyperlink));
            Assert.Equal("A0B0C0", reopened.GetThemeColor(
                PowerPointThemeColor.FollowedHyperlink));
            Assert.Equal("Aptos Display",
                reopened.GetThemeFonts().MajorLatin);
            Assert.Equal("Aptos", reopened.GetThemeFonts().MinorLatin);
            Assert.NotNull(reopened.OpenXmlDocument.PresentationPart!
                .SlideMasterParts.First().ThemePart?.Theme?.ThemeElements?
                .FormatScheme);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_RoundTripsSlideThemeOverrideAndColorMapping() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide(
                P.SlideLayoutValues.Blank);
            A.ThemeElements sourceElements = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.First().ThemePart!.Theme!
                .ThemeElements!;
            A.ColorScheme colors = (A.ColorScheme)sourceElements.ColorScheme!
                .CloneNode(true);
            A.Accent5Color accent5 = colors.GetFirstChild<A.Accent5Color>()!;
            accent5.RemoveAllChildren();
            accent5.Append(new A.RgbColorModelHex { Val = "ABCDEF" });
            ThemeOverridePart overridePart = slide.SlidePart
                .AddNewPart<ThemeOverridePart>();
            overridePart.ThemeOverride = new A.ThemeOverride(
                colors,
                sourceElements.FontScheme!.CloneNode(true),
                sourceElements.FormatScheme!.CloneNode(true));

            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptRoundTripTheme written = Assert.IsType<
                LegacyPptRoundTripTheme>(Assert.Single(
                    LegacyPptPresentation.Load(bytes).Slides).RoundTripTheme);

            Assert.True(written.IsOverride);
            Assert.Equal("ABCDEF",
                written.Colors[PowerPointThemeColor.Accent5]);
            Assert.Contains("clrMapOvr", written.ColorMappingXml);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(stream);
            A.ThemeOverride projected = Assert.IsType<A.ThemeOverride>(
                reopened.Slides[0].SlidePart.ThemeOverridePart?.ThemeOverride);
            Assert.Equal("ABCDEF", projected.ColorScheme!
                .GetFirstChild<A.Accent5Color>()!
                .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.NotNull(reopened.Slides[0].SlidePart.Slide!
                .ColorMapOverride?.GetFirstChild<A.MasterColorMapping>());
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void BinaryImport_ReportsMalformedRoundTripThemePackage() {
            LegacyPptPresentation source = LegacyPptPresentation.Load(
                AccessibilityFixture);
            byte[] document = source.Package.DocumentStream.ToArray();
            int recordOffset = FindRoundTripThemeRecord(document);
            Assert.True(recordOffset >= 0);
            document[recordOffset + 8] = (byte)'Q';
            byte[] malformed = source.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]> {
                    ["PowerPoint Document"] = document
                });

            LegacyPptPresentation decoded = LegacyPptPresentation.Load(
                malformed);

            Assert.Contains(decoded.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-ROUNDTRIP-THEME-INVALID");
            Assert.DoesNotContain(decoded.Masters, master =>
                master.RoundTripTheme?.ThemeXml != null);
        }

        [Fact]
        public void NativeWriter_RoundTripsNotesPageThemeOverride() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide(
                P.SlideLayoutValues.Blank);
            slide.Notes.Text = "Theme-aware speaker note";
            NotesSlidePart notesPart = slide.SlidePart.NotesSlidePart!;
            A.ThemeElements sourceElements = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.First().ThemePart!.Theme!
                .ThemeElements!;
            A.ColorScheme colors = (A.ColorScheme)sourceElements.ColorScheme!
                .CloneNode(true);
            A.Accent6Color accent6 = colors.GetFirstChild<A.Accent6Color>()!;
            accent6.RemoveAllChildren();
            accent6.Append(new A.RgbColorModelHex { Val = "C0FFEE" });
            ThemeOverridePart overridePart = notesPart
                .AddNewPart<ThemeOverridePart>();
            overridePart.ThemeOverride = new A.ThemeOverride(
                colors,
                sourceElements.FontScheme!.CloneNode(true),
                sourceElements.FormatScheme!.CloneNode(true));

            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptRoundTripTheme written = Assert.IsType<
                LegacyPptRoundTripTheme>(Assert.Single(
                    LegacyPptPresentation.Load(bytes).Slides).NotesPage!
                    .RoundTripTheme);

            Assert.True(written.IsOverride);
            Assert.Equal("C0FFEE",
                written.Colors[PowerPointThemeColor.Accent6]);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(stream);
            A.ThemeOverride projected = Assert.IsType<A.ThemeOverride>(
                reopened.Slides[0].SlidePart.NotesSlidePart?
                    .ThemeOverridePart?.ThemeOverride);
            Assert.Equal("C0FFEE", projected.ColorScheme!
                .GetFirstChild<A.Accent6Color>()!
                .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        private static int FindRoundTripThemeRecord(byte[] document) {
            for (int offset = 0; offset <= document.Length - 12; offset++) {
                if (document[offset] != 0 || document[offset + 1] != 0
                    || document[offset + 2] != 0x0E
                    || document[offset + 3] != 0x04) continue;
                int length = document[offset + 4]
                    | document[offset + 5] << 8
                    | document[offset + 6] << 16
                    | document[offset + 7] << 24;
                if (length > 4 && length <= document.Length - offset - 8
                    && document[offset + 8] == (byte)'P'
                    && document[offset + 9] == (byte)'K') return offset;
            }
            return -1;
        }
    }
}
