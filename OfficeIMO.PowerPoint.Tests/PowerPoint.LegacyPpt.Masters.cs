using System.Buffers.Binary;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using P = DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptMasterTests {
        private static string FixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "BasicPowerPoint.ppt");

        [Fact]
        public void DocumentAtomReader_DecodesCompleteDocumentAndMasterSettings() {
            var payload = new byte[40];
            BinaryPrimitives.WriteInt32LittleEndian(payload.AsSpan(0, 4), 7200);
            BinaryPrimitives.WriteInt32LittleEndian(payload.AsSpan(4, 4), 5400);
            BinaryPrimitives.WriteInt32LittleEndian(payload.AsSpan(8, 4), 5400);
            BinaryPrimitives.WriteInt32LittleEndian(payload.AsSpan(12, 4), 7200);
            BinaryPrimitives.WriteInt32LittleEndian(payload.AsSpan(16, 4), 3);
            BinaryPrimitives.WriteInt32LittleEndian(payload.AsSpan(20, 4), 2);
            BinaryPrimitives.WriteUInt32LittleEndian(payload.AsSpan(24, 4), 11);
            BinaryPrimitives.WriteUInt32LittleEndian(payload.AsSpan(28, 4), 12);
            BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(32, 2), 7);
            BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(34, 2),
                (ushort)LegacyPptSlideSizeType.A4Paper);
            payload[36] = 1;
            payload[37] = 1;
            payload[38] = 1;
            payload[39] = 1;
            var record = new LegacyPptRecord(payload, 0, 1, 0, 0x03E9,
                0, payload.Length);

            LegacyPptDocumentSettings settings = Assert.IsType<LegacyPptDocumentSettings>(
                LegacyPptDocumentAtomReader.Read(record));

            Assert.Equal(7200, settings.SlideWidth);
            Assert.Equal(5400, settings.SlideHeight);
            Assert.Equal(5400, settings.NotesWidth);
            Assert.Equal(7200, settings.NotesHeight);
            Assert.Equal(3, settings.ServerZoomNumerator);
            Assert.Equal(2, settings.ServerZoomDenominator);
            Assert.Equal(11U, settings.NotesMasterPersistId);
            Assert.Equal(12U, settings.HandoutMasterPersistId);
            Assert.Equal(7, settings.FirstSlideNumber);
            Assert.Equal(LegacyPptSlideSizeType.A4Paper, settings.SlideSizeType);
            Assert.True(settings.SaveWithFonts);
            Assert.True(settings.OmitTitlePlaceholders);
            Assert.True(settings.RightToLeft);
            Assert.True(settings.ShowComments);
            Assert.Null(LegacyPptDocumentAtomReader.Read(new LegacyPptRecord(
                new byte[39], 0, 1, 0, 0x03E9, 0, 39)));
        }

        [Fact]
        public void BinaryImport_ProjectsDocumentSettingsAndNotesMasterTopology() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(FixturePath);
            LegacyPptDocumentSettings settings = Assert.IsType<LegacyPptDocumentSettings>(
                legacy.DocumentSettings);
            Assert.Equal(settings.NotesMasterPersistId != 0, legacy.NotesMaster != null);
            if (legacy.NotesMaster != null) {
                Assert.Equal(LegacyPptSpecialMasterKind.Notes, legacy.NotesMaster.Kind);
                Assert.Equal(settings.NotesMasterPersistId, legacy.NotesMaster.PersistId);
                Assert.NotNull(legacy.NotesMaster.ColorScheme);
                Assert.NotEmpty(legacy.NotesMaster.Shapes);
            }

            using PowerPointPresentation projected = PowerPointPresentation.Load(FixturePath);
            P.Presentation root = projected.OpenXmlDocument.PresentationPart!.Presentation;
            Assert.Equal(settings.FirstSlideNumber, root.FirstSlideNum?.Value);
            Assert.Equal(checked((int)Math.Round(100000D
                    * settings.ServerZoomNumerator / settings.ServerZoomDenominator,
                MidpointRounding.AwayFromZero)), root.ServerZoom?.Value);
            Assert.Equal(!settings.OmitTitlePlaceholders,
                root.ShowSpecialPlaceholderOnTitleSlide?.Value);
            Assert.Equal(settings.RightToLeft, root.RightToLeft?.Value);
            Assert.Equal(settings.SaveWithFonts, root.EmbedTrueTypeFonts?.Value);
            Assert.Equal(settings.ShowComments, projected.OpenXmlDocument.PresentationPart!
                .ViewPropertiesPart!.ViewProperties!.ShowComments?.Value);
            Assert.Equal(ToEmus(settings.NotesWidth), root.NotesSize!.Cx!.Value);
            Assert.Equal(ToEmus(settings.NotesHeight), root.NotesSize.Cy!.Value);

            if (legacy.NotesMaster != null) {
                NotesMasterPart notesPart = projected.OpenXmlDocument.PresentationPart!
                    .NotesMasterPart!;
                Assert.Equal("Binary Notes Master",
                    notesPart.NotesMaster!.CommonSlideData!.Name!.Value);
                Assert.NotNull(notesPart.ThemePart?.Theme?.ThemeElements?.ColorScheme);
                Assert.NotEmpty(notesPart.NotesMaster.CommonSlideData.ShapeTree!
                    .Descendants<P.PlaceholderShape>());
            }
            Assert.Empty(projected.ValidateDocument());

            LegacyPptImportReport report = legacy.CreateImportReport();
            Assert.Equal(legacy.NotesMaster == null ? 0 : 1, report.SpecialMasterCount);
            Assert.Equal(legacy.NotesMaster?.Shapes.Count ?? 0,
                report.SpecialMasterShapeCount);
        }

        [Fact]
        public void NativeWriter_RoundTripsDocumentPageAndDisplaySettings() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.SlideSize.SetSizeEmus(9906000, 6858000, P.SlideSizeValues.A4);
            P.Presentation root = presentation.OpenXmlDocument.PresentationPart!.Presentation;
            root.NotesSize = new P.NotesSize { Cx = 6858000, Cy = 9906000 };
            root.FirstSlideNum = 7;
            root.ShowSpecialPlaceholderOnTitleSlide = false;
            root.RightToLeft = true;
            root.EmbedTrueTypeFonts = true;
            root.ServerZoom = 75000;
            presentation.OpenXmlDocument.PresentationPart!.ViewPropertiesPart!
                .ViewProperties!.ShowComments = true;
            presentation.AddSlide();

            LegacyPptDocumentSettings settings = Assert.IsType<LegacyPptDocumentSettings>(
                LegacyPptPresentation.Load(presentation.ToBytes(PowerPointFileFormat.Ppt))
                    .DocumentSettings);

            Assert.Equal(LegacyPptSlideSizeType.A4Paper, settings.SlideSizeType);
            Assert.Equal(7, settings.FirstSlideNumber);
            Assert.True(settings.OmitTitlePlaceholders);
            Assert.True(settings.RightToLeft);
            Assert.True(settings.SaveWithFonts);
            Assert.True(settings.ShowComments);
            Assert.Equal(3, settings.ServerZoomNumerator);
            Assert.Equal(4, settings.ServerZoomDenominator);
            Assert.Equal(6858000, ToEmus(settings.NotesWidth));
            Assert.Equal(9906000, ToEmus(settings.NotesHeight));
        }

        [Fact]
        public void OfficeArtStyleDecoder_ExposesBackgroundFillInputs() {
            OfficeArtShapeStyle style = OfficeArtShapeStyle.Decode(new[] {
                new OfficeArtProperty(0, 0x0180, 4),
                new OfficeArtProperty(1, 0x0181, 0x00030201),
                new OfficeArtProperty(2, 0x0182, 0x00008000),
                new OfficeArtProperty(3, 0x0183, 0x00060504),
                new OfficeArtProperty(4, 0x0184, 0x00004000),
                new OfficeArtProperty(5, 0x4186, 3),
                new OfficeArtProperty(6, 0x018B, unchecked((uint)(-45 * 65536))),
                new OfficeArtProperty(7, 0x018C, 25)
            });

            Assert.Equal(4U, style.FillType);
            Assert.NotNull(style.FillColor);
            Assert.Equal(0.5D, style.FillOpacity);
            Assert.NotNull(style.FillBackColor);
            Assert.Equal(0.25D, style.FillBackOpacity);
            Assert.Equal(3, style.FillBlipStoreIndex);
            Assert.Equal(-45D, style.FillAngleDegrees);
            Assert.Equal(25, style.FillFocusPercent);
        }

        [Fact]
        public void OfficeArtStyleDecoder_DecodesAndRejectsGradientStopArrays() {
            byte[] valid = CreateGradientStopArray(
                (0x00030201U, 0x00000000U),
                (0x00060504U, 0x00008000U),
                (0x00090807U, 0x00010000U));
            OfficeArtShapeStyle decoded = OfficeArtShapeStyle.Decode(new[] {
                new OfficeArtProperty(0, 0x8197, checked((uint)valid.Length),
                    valid.Length, complexData: valid)
            });

            Assert.False(decoded.IsFillGradientStopTableTruncated);
            Assert.Collection(decoded.FillGradientStops,
                stop => Assert.Equal(0D, stop.Position),
                stop => Assert.Equal(0.5D, stop.Position),
                stop => Assert.Equal(1D, stop.Position));
            Assert.Equal(0x00060504U, decoded.FillGradientStops[1].Color.Value);

            byte[] descending = CreateGradientStopArray(
                (0x00030201U, 0x00010000U),
                (0x00060504U, 0x00008000U));
            OfficeArtShapeStyle rejected = OfficeArtShapeStyle.Decode(new[] {
                new OfficeArtProperty(0, 0x8197, checked((uint)descending.Length),
                    descending.Length, complexData: descending)
            });

            Assert.True(rejected.IsFillGradientStopTableTruncated);
            Assert.Empty(rejected.FillGradientStops);
        }

        [Fact]
        public void BinaryImport_DecodesAndProjectsOfficeArtMasterBackground() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(FixturePath);
            LegacyPptMaster mainMaster = legacy.Masters.First(master => master.IsMainMaster);
            LegacyPptBackground background = Assert.IsType<LegacyPptBackground>(
                mainMaster.Background);

            Assert.True(background.HasProjectableFill);
            Assert.NotEqual(LegacyPptBackgroundKind.Unsupported, background.Kind);
            LegacyPptImportReport report = legacy.CreateImportReport();
            Assert.True(report.BackgroundCount > 0);
            Assert.True(report.ProjectableBackgroundCount > 0);

            using PowerPointPresentation projected = PowerPointPresentation.Load(FixturePath);
            P.Background projectedBackground = Assert.IsType<P.Background>(projected
                .OpenXmlDocument.PresentationPart!.SlideMasterParts.First().SlideMaster!
                .CommonSlideData!.Background);
            Assert.NotNull(projectedBackground.BackgroundProperties);
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void BinaryCorpus_BackgroundShapesRemainTypedAndProjectable() {
            string corpus = Path.Combine(AppContext.BaseDirectory, "Documents",
                "LegacyPptCorpus");
            foreach (string path in Directory.GetFiles(corpus, "*.ppt")) {
                LegacyPptPresentation legacy = LegacyPptPresentation.Load(path);
                LegacyPptBackground[] backgrounds = legacy.Slides
                    .Select(slide => slide.Background)
                    .Concat(legacy.Masters.Select(master => master.Background))
                    .Concat(new[] {
                        legacy.NotesMaster?.Background,
                        legacy.HandoutMaster?.Background
                    })
                    .Where(background => background != null)
                    .Cast<LegacyPptBackground>()
                    .ToArray();

                Assert.NotEmpty(backgrounds);
                Assert.DoesNotContain(backgrounds,
                    background => background.Kind == LegacyPptBackgroundKind.Unsupported);
                Assert.All(backgrounds,
                    background => Assert.True(background.HasProjectableFill));
                Assert.DoesNotContain(legacy.Diagnostics,
                    diagnostic => diagnostic.Code == "PPT-BACKGROUND-PARTIAL");
            }
        }

        private static int ToEmus(int masterUnits) =>
            checked((int)Math.Round(masterUnits * 1587.5d, MidpointRounding.AwayFromZero));

        private static byte[] CreateGradientStopArray(
            params (uint Color, uint Position)[] stops) {
            var data = new byte[checked(6 + stops.Length * 8)];
            BinaryPrimitives.WriteUInt16LittleEndian(data.AsSpan(0, 2),
                checked((ushort)stops.Length));
            BinaryPrimitives.WriteUInt16LittleEndian(data.AsSpan(2, 2),
                checked((ushort)stops.Length));
            BinaryPrimitives.WriteUInt16LittleEndian(data.AsSpan(4, 2), 8);
            for (int index = 0; index < stops.Length; index++) {
                int offset = 6 + index * 8;
                BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(offset, 4),
                    stops[index].Color);
                BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(offset + 4, 4),
                    stops[index].Position);
            }
            return data;
        }
    }
}
