using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
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
        public void ImportedMasterThemeEdit_AppendsPreservingIncrementalRecord() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                AccessibilityFixture);
            LegacyPptMaster[] originalMainMasters = original.Masters
                .Where(master => master.IsMainMaster).ToArray();
            int masterIndex = Array.FindIndex(originalMainMasters,
                master => master.RoundTripTheme?.ThemeXml != null);
            Assert.True(masterIndex >= 0);
            LegacyPptMaster originalMaster = originalMainMasters[masterIndex];

            using PowerPointPresentation imported = PowerPointPresentation.Load(
                AccessibilityFixture);
            imported.SetThemeColor(PowerPointThemeColor.Accent5,
                "123456", masterIndex);
            imported.SetThemeColor(PowerPointThemeColor.Accent1,
                "654321", masterIndex);
            imported.SetThemeLatinFonts("Aptos Display", "Aptos",
                masterIndex);

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] bytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(bytes);
            LegacyPptMaster savedMaster = Assert.Single(saved.Masters,
                master => master.MasterId == originalMaster.MasterId);

            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
            Assert.Equal("123456", savedMaster.RoundTripTheme?
                .Colors[PowerPointThemeColor.Accent5]);
            Assert.Equal("654321", savedMaster.RoundTripTheme?
                .Colors[PowerPointThemeColor.Accent1]);
            Assert.Equal("654321", savedMaster.ColorScheme?.Accent1);
            Assert.Equal(originalMaster.ColorScheme?.Background,
                savedMaster.ColorScheme?.Background);
            Assert.Equal(originalMaster.ColorScheme?.Text,
                savedMaster.ColorScheme?.Text);
            Assert.Equal(originalMaster.ColorScheme?.Shadow,
                savedMaster.ColorScheme?.Shadow);
            Assert.Equal(originalMaster.ColorScheme?.TitleText,
                savedMaster.ColorScheme?.TitleText);
            Assert.Equal(originalMaster.ColorScheme?.Fill,
                savedMaster.ColorScheme?.Fill);
            Assert.Equal(originalMaster.ColorScheme?.Accent2,
                savedMaster.ColorScheme?.Accent2);
            Assert.Equal(originalMaster.ColorScheme?.Accent3,
                savedMaster.ColorScheme?.Accent3);
            Assert.Equal("Aptos Display",
                savedMaster.RoundTripTheme?.MajorLatinFont);
            Assert.Equal("Aptos",
                savedMaster.RoundTripTheme?.MinorLatinFont);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(stream);
            Assert.Equal("123456", reopened.GetThemeColor(
                PowerPointThemeColor.Accent5, masterIndex));
            Assert.Equal("654321", reopened.GetThemeColor(
                PowerPointThemeColor.Accent1, masterIndex));
            Assert.Equal("Aptos Display",
                reopened.GetThemeFonts(masterIndex).MajorLatin);
            Assert.Equal("Aptos",
                reopened.GetThemeFonts(masterIndex).MinorLatin);
            Assert.Empty(reopened.ValidateDocument());

            IReadOnlyList<byte[]> originalUnrelated =
                ReadUnrelatedThemeChildren(original, originalMaster.PersistId);
            IReadOnlyList<byte[]> savedUnrelated =
                ReadUnrelatedThemeChildren(saved, savedMaster.PersistId);
            Assert.Equal(originalUnrelated.Count, savedUnrelated.Count);
            for (int index = 0; index < originalUnrelated.Count; index++) {
                Assert.True(originalUnrelated[index]
                    .SequenceEqual(savedUnrelated[index]));
            }
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
            byte[] bytes = CreateSlideThemeOverrideBytes(
                accent5: "ABCDEF", accent1: "123456");
            LegacyPptSlide writtenSlide = Assert.Single(
                LegacyPptPresentation.Load(bytes).Slides);
            LegacyPptRoundTripTheme written = Assert.IsType<
                LegacyPptRoundTripTheme>(writtenSlide.RoundTripTheme);

            Assert.True(written.IsOverride);
            Assert.Equal("ABCDEF",
                written.Colors[PowerPointThemeColor.Accent5]);
            Assert.Equal("123456", writtenSlide.ColorScheme?.Accent1);
            Assert.False(writtenSlide.FollowsMasterColorScheme);
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
        public void ImportedSlideThemeEdit_AppendsPreservingIncrementalRecord() {
            byte[] sourceBytes = CreateSlideThemeOverrideBytes(
                accent5: "ABCDEF", accent1: "123456");
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);
            LegacyPptSlide originalSlide = Assert.Single(original.Slides);

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            A.ThemeOverride theme = Assert.IsType<A.ThemeOverride>(imported
                .Slides[0].SlidePart.ThemeOverridePart?.ThemeOverride);
            SetThemeOverrideColor<A.Accent5Color>(theme, "5A6B7C");
            SetThemeOverrideColor<A.Accent1Color>(theme, "102938");

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptSlide savedSlide = Assert.Single(saved.Slides);

            Assert.Equal("5A6B7C", savedSlide.RoundTripTheme?
                .Colors[PowerPointThemeColor.Accent5]);
            Assert.Equal("102938", savedSlide.RoundTripTheme?
                .Colors[PowerPointThemeColor.Accent1]);
            Assert.Equal("102938", savedSlide.ColorScheme?.Accent1);
            Assert.False(savedSlide.FollowsMasterColorScheme);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
            IReadOnlyList<byte[]> originalUnrelated =
                ReadUnrelatedThemeChildren(original, originalSlide.PersistId);
            IReadOnlyList<byte[]> savedUnrelated =
                ReadUnrelatedThemeChildren(saved, savedSlide.PersistId);
            Assert.Equal(originalUnrelated.Count, savedUnrelated.Count);
            for (int index = 0; index < originalUnrelated.Count; index++) {
                Assert.True(originalUnrelated[index]
                    .SequenceEqual(savedUnrelated[index]));
            }

            using var reopenedInput = new MemoryStream(savedBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            A.ThemeOverride reopenedTheme = Assert.IsType<A.ThemeOverride>(
                reopened.Slides[0].SlidePart.ThemeOverridePart?.ThemeOverride);
            Assert.Equal("5A6B7C", reopenedTheme.ColorScheme!
                .GetFirstChild<A.Accent5Color>()!
                .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal("102938", reopenedTheme.ColorScheme!
                .GetFirstChild<A.Accent1Color>()!
                .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedLayoutThemeEdit_MaterializesIntoAffectedSlides() {
            byte[] sourceBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                created.AddSlide(P.SlideLayoutValues.Blank);
                created.AddSlide(P.SlideLayoutValues.Blank);
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            SlideLayoutPart layoutPart = imported.Slides[0].SlidePart
                .SlideLayoutPart!;
            Assert.All(imported.Slides, slide => Assert.Same(layoutPart,
                slide.SlidePart.SlideLayoutPart));
            Assert.True(imported.LegacyPptProjectionMap!
                .IsEditableProjectedLayoutThemePart(
                    layoutPart.Uri.ToString()));
            A.ThemeElements sourceElements = layoutPart.SlideMasterPart!
                .ThemePart!.Theme!.ThemeElements!;
            A.ColorScheme colors = (A.ColorScheme)sourceElements.ColorScheme!
                .CloneNode(true);
            SetThemeColor<A.Accent5Color>(colors, "5A6B7C");
            SetThemeColor<A.Accent1Color>(colors, "102938");
            ThemeOverridePart overridePart = layoutPart
                .AddNewPart<ThemeOverridePart>();
            overridePart.ThemeOverride = new A.ThemeOverride(
                colors,
                sourceElements.FontScheme!.CloneNode(true),
                sourceElements.FormatScheme!.CloneNode(true));

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);

            Assert.Equal(2, saved.Slides.Count);
            Assert.All(saved.Slides, slide => {
                Assert.Equal("5A6B7C", slide.RoundTripTheme?
                    .Colors[PowerPointThemeColor.Accent5]);
                Assert.Equal("102938", slide.ColorScheme?.Accent1);
                Assert.False(slide.FollowsMasterColorScheme);
            });
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            using var reopenedInput = new MemoryStream(savedBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            Assert.All(reopened.Slides, slide => {
                A.ThemeOverride theme = Assert.IsType<A.ThemeOverride>(
                    slide.SlidePart.ThemeOverridePart?.ThemeOverride);
                Assert.Equal("5A6B7C", theme.ColorScheme!
                    .GetFirstChild<A.Accent5Color>()!
                    .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
                Assert.Equal("102938", theme.ColorScheme!
                    .GetFirstChild<A.Accent1Color>()!
                    .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            });
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
        public void BinaryImport_RejectsRoundTripThemeEntryCountBeforeOpeningEntries() {
            LegacyPptPresentation source = LegacyPptPresentation.Load(
                AccessibilityFixture);
            byte[] document = source.Package.DocumentStream.ToArray();
            int recordOffset = FindRoundTripThemeRecord(document);
            Assert.True(recordOffset >= 0);
            int payloadLength = ReadInt32(document, recordOffset + 4);
            byte[] oversizedDirectory = CreateThemeArchive(
                entryCount: 65, payloadLength);
            Buffer.BlockCopy(oversizedDirectory, 0, document,
                recordOffset + 8, payloadLength);
            byte[] guarded = source.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]> {
                    ["PowerPoint Document"] = document
                });

            LegacyPptPresentation decoded = LegacyPptPresentation.Load(
                guarded);

            Assert.Contains(decoded.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-ROUNDTRIP-THEME-INVALID"
                && diagnostic.Message.Contains("too many entries",
                    StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void NativeWriter_RoundTripsNotesPageThemeOverride() {
            byte[] bytes = CreateNotesThemeOverrideBytes(
                accent6: "C0FFEE", accent1: "123456",
                background: "203040", text: "Theme-aware speaker note");
            LegacyPptNotesPage notesPage = Assert.IsType<LegacyPptNotesPage>(
                Assert.Single(LegacyPptPresentation.Load(bytes).Slides)
                    .NotesPage);
            LegacyPptRoundTripTheme written = Assert.IsType<
                LegacyPptRoundTripTheme>(notesPage.RoundTripTheme);

            Assert.True(written.IsOverride);
            Assert.Equal("C0FFEE",
                written.Colors[PowerPointThemeColor.Accent6]);
            Assert.Equal("123456", notesPage.ColorScheme?.Accent1);
            Assert.False(notesPage.FollowsMasterColorScheme);
            Assert.False(notesPage.FollowsMasterBackground);
            Assert.Equal("203040", Assert.IsType<LegacyPptBackground>(
                notesPage.Background).ForegroundColor);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(stream);
            A.ThemeOverride projected = Assert.IsType<A.ThemeOverride>(
                reopened.Slides[0].SlidePart.NotesSlidePart?
                    .ThemeOverridePart?.ThemeOverride);
            Assert.Equal("C0FFEE", projected.ColorScheme!
                .GetFirstChild<A.Accent6Color>()!
                .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal("203040", reopened.Slides[0].SlidePart.NotesSlidePart!
                .NotesSlide!.CommonSlideData!.Background!
                .BackgroundProperties!.GetFirstChild<A.SolidFill>()!
                .RgbColorModelHex!.Val!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedNotesThemeBackgroundAndTextEdit_AppendsPreservingRecord() {
            byte[] sourceBytes = CreateNotesThemeOverrideBytes(
                accent6: "C0FFEE", accent1: "123456",
                background: "203040", text: "Original speaker note");
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);
            LegacyPptNotesPage originalNotes = Assert.IsType<
                LegacyPptNotesPage>(Assert.Single(original.Slides).NotesPage);

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            PowerPointSlide slide = imported.Slides[0];
            slide.Notes.Text = "Edited speaker note";
            NotesSlidePart notesPart = slide.SlidePart.NotesSlidePart!;
            A.ThemeOverride theme = Assert.IsType<A.ThemeOverride>(
                notesPart.ThemeOverridePart?.ThemeOverride);
            SetThemeOverrideColor<A.Accent6Color>(theme, "6A7B8C");
            SetThemeOverrideColor<A.Accent1Color>(theme, "102938");
            notesPart.NotesSlide!.CommonSlideData!.Background = new P.Background(
                new P.BackgroundProperties(
                    new A.SolidFill(
                        new A.RgbColorModelHex { Val = "405060" })));

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptNotesPage savedNotes = Assert.IsType<LegacyPptNotesPage>(
                Assert.Single(saved.Slides).NotesPage);

            Assert.Equal("Edited speaker note", savedNotes.Text);
            Assert.Equal("6A7B8C", savedNotes.RoundTripTheme?
                .Colors[PowerPointThemeColor.Accent6]);
            Assert.Equal("102938", savedNotes.ColorScheme?.Accent1);
            Assert.Equal("405060", Assert.IsType<LegacyPptBackground>(
                savedNotes.Background).ForegroundColor);
            Assert.False(savedNotes.FollowsMasterColorScheme);
            Assert.False(savedNotes.FollowsMasterBackground);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
            IReadOnlyList<byte[]> originalUnrelated =
                ReadUnrelatedNotesChildren(original, originalNotes.PersistId);
            IReadOnlyList<byte[]> savedUnrelated =
                ReadUnrelatedNotesChildren(saved, savedNotes.PersistId);
            Assert.Equal(originalUnrelated.Count, savedUnrelated.Count);
            for (int index = 0; index < originalUnrelated.Count; index++) {
                Assert.True(originalUnrelated[index]
                    .SequenceEqual(savedUnrelated[index]));
            }

            using var reopenedInput = new MemoryStream(savedBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            Assert.Equal("Edited speaker note", reopened.Slides[0].Notes.Text);
            NotesSlidePart reopenedNotesPart = reopened.Slides[0].SlidePart
                .NotesSlidePart!;
            Assert.Equal("6A7B8C", reopenedNotesPart.ThemeOverridePart!
                .ThemeOverride!.ColorScheme!
                .GetFirstChild<A.Accent6Color>()!
                .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal("405060", reopenedNotesPart.NotesSlide!
                .CommonSlideData!.Background!.BackgroundProperties!
                .GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        private static byte[] CreateSlideThemeOverrideBytes(string accent5,
            string accent1) {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide(
                P.SlideLayoutValues.Blank);
            A.ThemeElements sourceElements = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.First().ThemePart!.Theme!
                .ThemeElements!;
            A.ColorScheme colors = (A.ColorScheme)sourceElements.ColorScheme!
                .CloneNode(true);
            SetThemeColor<A.Accent5Color>(colors, accent5);
            SetThemeColor<A.Accent1Color>(colors, accent1);
            ThemeOverridePart overridePart = slide.SlidePart
                .AddNewPart<ThemeOverridePart>();
            overridePart.ThemeOverride = new A.ThemeOverride(
                colors,
                sourceElements.FontScheme!.CloneNode(true),
                sourceElements.FormatScheme!.CloneNode(true));
            return presentation.ToBytes(PowerPointFileFormat.Ppt);
        }

        private static byte[] CreateNotesThemeOverrideBytes(string accent6,
            string accent1, string background, string text) {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide(
                P.SlideLayoutValues.Blank);
            slide.Notes.Text = text;
            NotesSlidePart notesPart = slide.SlidePart.NotesSlidePart!;
            A.ThemeElements sourceElements = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.First().ThemePart!.Theme!
                .ThemeElements!;
            A.ColorScheme colors = (A.ColorScheme)sourceElements.ColorScheme!
                .CloneNode(true);
            SetThemeColor<A.Accent6Color>(colors, accent6);
            SetThemeColor<A.Accent1Color>(colors, accent1);
            ThemeOverridePart overridePart = notesPart
                .AddNewPart<ThemeOverridePart>();
            overridePart.ThemeOverride = new A.ThemeOverride(
                colors,
                sourceElements.FontScheme!.CloneNode(true),
                sourceElements.FormatScheme!.CloneNode(true));
            notesPart.NotesSlide!.CommonSlideData!.Background =
                new P.Background(new P.BackgroundProperties(
                    new A.SolidFill(
                        new A.RgbColorModelHex { Val = background })));
            return presentation.ToBytes(PowerPointFileFormat.Ppt);
        }

        private static void SetThemeOverrideColor<T>(A.ThemeOverride theme,
            string value) where T : OpenXmlCompositeElement {
            T color = Assert.IsType<T>(theme.ColorScheme?
                .GetFirstChild<T>());
            color.RemoveAllChildren();
            color.Append(new A.RgbColorModelHex { Val = value });
        }

        private static void SetThemeColor<T>(A.ColorScheme theme,
            string value) where T : OpenXmlCompositeElement {
            T color = Assert.IsType<T>(theme.GetFirstChild<T>());
            color.RemoveAllChildren();
            color.Append(new A.RgbColorModelHex { Val = value });
        }

        private static IReadOnlyList<byte[]> ReadUnrelatedThemeChildren(
            LegacyPptPresentation presentation, uint persistId) {
            LegacyPptPersistObject persistObject =
                presentation.Package.PersistObjects[persistId];
            LegacyPptRecord record = LegacyPptRecordReader.ReadSingle(
                persistObject.RecordBytes, 0, new LegacyPptImportOptions());
            return record.Children.Where(child =>
                    child.Type is not 0x040E and not 0x040F
                    && !(child.Type == 0x07F0 && child.Instance == 1))
                .Select(child => child.CopyRecordBytes()).ToArray();
        }

        private static IReadOnlyList<byte[]> ReadUnrelatedNotesChildren(
            LegacyPptPresentation presentation, uint persistId) {
            LegacyPptPersistObject persistObject =
                presentation.Package.PersistObjects[persistId];
            LegacyPptRecord record = LegacyPptRecordReader.ReadSingle(
                persistObject.RecordBytes, 0, new LegacyPptImportOptions());
            return record.Children.Where(child => child.Type != 0x040C
                    && child.Type is not 0x040E and not 0x040F
                    && !(child.Type == 0x07F0 && child.Instance == 1))
                .Select(child => child.CopyRecordBytes()).ToArray();
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

        private static byte[] CreateThemeArchive(int entryCount,
            int targetLength) {
            const uint LocalFileHeaderSignature = 0x04034B50U;
            const uint CentralDirectoryHeaderSignature = 0x02014B50U;
            const uint EndOfCentralDirectorySignature = 0x06054B50U;
            using var output = new MemoryStream();
            using (var writer = new BinaryWriter(output,
                       Encoding.UTF8, leaveOpen: true)) {
                writer.Write(LocalFileHeaderSignature);
                writer.Write(new byte[26]);
                int centralDirectoryOffset = checked((int)output.Position);
                for (int index = 0; index < entryCount; index++) {
                    writer.Write(CentralDirectoryHeaderSignature);
                    writer.Write(new byte[42]);
                }
                int centralDirectorySize = checked((int)output.Position)
                    - centralDirectoryOffset;
                writer.Write(EndOfCentralDirectorySignature);
                writer.Write((ushort)0);
                writer.Write((ushort)0);
                writer.Write(checked((ushort)entryCount));
                writer.Write(checked((ushort)entryCount));
                writer.Write(checked((uint)centralDirectorySize));
                writer.Write(checked((uint)centralDirectoryOffset));
                writer.Write((ushort)0);
            }
            byte[] bytes = output.ToArray();
            int commentLength = targetLength - bytes.Length;
            Assert.InRange(commentLength, 0, ushort.MaxValue);
            int endOfCentralDirectory = bytes.Length - 22;
            Array.Resize(ref bytes, targetLength);
            bytes[endOfCentralDirectory + 20] = (byte)commentLength;
            bytes[endOfCentralDirectory + 21] =
                (byte)(commentLength >> 8);
            return bytes;
        }

        private static int ReadInt32(byte[] bytes, int offset) =>
            bytes[offset]
            | bytes[offset + 1] << 8
            | bytes[offset + 2] << 16
            | bytes[offset + 3] << 24;
    }
}
