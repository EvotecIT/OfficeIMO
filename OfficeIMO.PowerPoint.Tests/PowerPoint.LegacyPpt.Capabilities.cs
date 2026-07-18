using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptCapabilityTests {
        [Fact]
        public void CapabilityContract_CoversEveryFeatureExactlyOnce() {
            LegacyPptFeature[] features = Enum.GetValues(typeof(LegacyPptFeature)).Cast<LegacyPptFeature>().ToArray();

            Assert.Equal(features.Length, LegacyPptCapabilityCatalog.Capabilities.Count);
            Assert.Equal(features.Length, LegacyPptCapabilityCatalog.Capabilities.Select(row => row.Feature).Distinct().Count());
            foreach (LegacyPptFeature feature in features) {
                Assert.Equal(feature, LegacyPptCapabilityCatalog.Get(feature).Feature);
            }
        }

        [Fact]
        public void CapabilityContract_SerializesDeterministicallyWithoutProvisionalStates() {
            string first = LegacyPptCapabilityCatalog.ToJson();
            string second = LegacyPptCapabilityCatalog.ToJson();

            Assert.Equal(first, second);
            Assert.Contains("\"schemaVersion\":1", first);
            Assert.Contains("\"feature\":\"UnknownRecordsAndStreams\"", first);
            Assert.False(LegacyPptCapabilityCatalog.HasRemainingParityWork);
            Assert.Empty(LegacyPptCapabilityCatalog.RemainingParityWork);
            Assert.DoesNotContain("Planned", first);
            Assert.Contains("| Preservation | UnknownRecordsAndStreams |",
                LegacyPptCapabilityCatalog.ToMarkdown());
        }

        [Fact]
        public void WritePreflight_FindingsReferenceNonNativeCapabilityRows() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.Hidden = true;
            PowerPointAutoShape shape = slide.AddRectangle(
                100000, 100000, 1000000, 500000);
            shape.Fill("FF0000");
            shape.SetGlow("4472C4", radiusPoints: 4D);

            LegacyPptWritePreflightReport report = presentation.AnalyzeLegacyPptWrite();

            Assert.NotEmpty(report.Findings);
            foreach (LegacyPptWriteFinding finding in report.Findings) {
                LegacyPptCapability capability = LegacyPptCapabilityCatalog.Get(finding.Feature);
                Assert.NotEqual(LegacyPptCapabilityState.Native, capability.PptxToBinary);
            }
        }

        [Fact]
        public void CapabilityContract_ReportsNativeRasterAuthoringAndPreservingRoundTrip() {
            LegacyPptCapability raster = LegacyPptCapabilityCatalog.Get(LegacyPptFeature.RasterPictures);

            Assert.Equal(LegacyPptCapabilityState.Native, raster.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Preserved, raster.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native, raster.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Native, raster.PptxToBinary);
            Assert.Contains("deduplicates identical payloads", raster.Note);
            Assert.Contains("main, notes, and handout masters", raster.Note);
            Assert.Contains("unused picture-bearing layouts are loss-blocked",
                raster.Note);
            Assert.Contains("protection-lock", raster.Note);

            LegacyPptCapability crop = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.PictureCrop);
            Assert.Equal(LegacyPptCapabilityState.Native,
                crop.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                crop.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                crop.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native,
                crop.PptxToBinary);
        }

        [Fact]
        public void CapabilityContract_ReportsNativeMetafileAuthoring() {
            LegacyPptCapability metafiles = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.MetafilePictures);

            Assert.Equal(LegacyPptRepresentability.Native,
                metafiles.Representability);
            Assert.Equal(LegacyPptCapabilityState.Native,
                metafiles.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                metafiles.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                metafiles.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native,
                metafiles.PptxToBinary);
            Assert.Contains("placeable-WMF", metafiles.Note);
        }

        [Fact]
        public void CapabilityContract_ReportsNativeClassicShapeAuthoringWithExplicitGaps() {
            LegacyPptCapability autoShapes = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.AutoShapes);
            LegacyPptCapability connectors = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.Connectors);
            LegacyPptCapability groups = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.Groups);
            LegacyPptCapability transforms = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.ShapeTransforms);
            LegacyPptCapability styles = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.ShapeStyles);
            LegacyPptCapability effects = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.ShapeEffects);

            Assert.Equal(LegacyPptCapabilityState.Native,
                autoShapes.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                autoShapes.PptxToBinary);
            Assert.Contains("canonical classic preset families",
                autoShapes.Note);
            Assert.Equal(LegacyPptCapabilityState.Native,
                connectors.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                connectors.PptxToBinary);
            Assert.Contains("Fresh or edited attachment rules",
                connectors.Note);
            Assert.Equal(LegacyPptRepresentability.Native,
                groups.Representability);
            Assert.Equal(LegacyPptCapabilityState.Native,
                groups.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                groups.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                groups.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native,
                groups.PptxToBinary);
            Assert.Contains("child anchors", groups.Note);
            Assert.Contains("Imported reparenting", groups.Note);
            Assert.Equal(LegacyPptCapabilityState.Native,
                transforms.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Native,
                transforms.PptxToBinary);
            Assert.Contains("child transform edits",
                transforms.Note);
            Assert.Equal(LegacyPptCapabilityState.Native,
                styles.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                styles.PptxToBinary);
            Assert.Contains("per-shape hidden state", styles.Note);
            Assert.Equal(LegacyPptCapabilityState.Native,
                effects.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                effects.PptxToBinary);
            Assert.Contains("one RGB outer shadow", effects.Note);
        }

        [Fact]
        public void CapabilityContract_ReportsExplicitStaticChartConversion() {
            LegacyPptCapability charts = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.Charts);

            Assert.Equal(LegacyPptRepresentability.Approximation,
                charts.Representability);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                charts.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Converted,
                charts.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                charts.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Converted,
                charts.PptxToBinary);
            Assert.Contains("PPT-WRITE-CHART-CONVERTED", charts.Note);
        }

        [Fact]
        public void CapabilityContract_ReportsNativeEditableTables() {
            LegacyPptCapability tables = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.Tables);

            Assert.Equal(LegacyPptRepresentability.Native,
                tables.Representability);
            Assert.Equal(LegacyPptCapabilityState.Native,
                tables.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                tables.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Native,
                tables.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native,
                tables.PptxToBinary);
            Assert.Contains("native editable OfficeArt table groups", tables.Note);
        }

        [Fact]
        public void CapabilityContract_ReportsExplicitStaticSmartArtConversion() {
            LegacyPptCapability smartArt = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.SmartArt);

            Assert.Equal(LegacyPptRepresentability.Approximation,
                smartArt.Representability);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                smartArt.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Converted,
                smartArt.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                smartArt.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Converted,
                smartArt.PptxToBinary);
            Assert.Contains("PPT-WRITE-SMARTART-CONVERTED", smartArt.Note);
        }

        [Fact]
        public void CapabilityContract_ReportsAccessibilityImportAndPreservingRoundTrip() {
            LegacyPptCapability accessibility = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.AccessibilityMetadata);

            Assert.Equal(LegacyPptCapabilityState.Native, accessibility.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native, accessibility.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native, accessibility.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Native, accessibility.PptxToBinary);
            Assert.Contains("explicitly loss-blocked", accessibility.Note);
        }

        [Fact]
        public void CapabilityContract_ReportsCompletedTextLanes() {
            foreach (LegacyPptFeature feature in new[] {
                         LegacyPptFeature.PlainText,
                         LegacyPptFeature.BulletsAndNumbering
                     }) {
                LegacyPptCapability capability =
                    LegacyPptCapabilityCatalog.Get(feature);
                Assert.Equal(LegacyPptCapabilityState.Native,
                    capability.ImportToEditableModel);
                Assert.Equal(LegacyPptCapabilityState.Native,
                    capability.NewBinaryWrite);
                Assert.Equal(LegacyPptCapabilityState.Native,
                    capability.BinaryRoundTrip);
                Assert.Equal(LegacyPptCapabilityState.Native,
                    capability.PptxToBinary);
            }

            foreach (LegacyPptFeature feature in new[] {
                         LegacyPptFeature.RichText,
                         LegacyPptFeature.ParagraphFormatting
                     }) {
                LegacyPptCapability capability =
                    LegacyPptCapabilityCatalog.Get(feature);
                Assert.Equal(LegacyPptCapabilityState.Native,
                    capability.ImportToEditableModel);
                Assert.Equal(LegacyPptCapabilityState.Native,
                    capability.NewBinaryWrite);
                Assert.Equal(LegacyPptCapabilityState.Preserved,
                    capability.BinaryRoundTrip);
                Assert.Equal(LegacyPptCapabilityState.Native,
                    capability.PptxToBinary);
            }
            LegacyPptCapability richText = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.RichText);
            Assert.Contains("primary and alternate language LCIDs",
                richText.Note);
            Assert.Contains("paragraph-end markers", richText.Note);

            LegacyPptCapability autoFit = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.TextAutoFit);
            Assert.Equal(LegacyPptCapabilityState.Native,
                autoFit.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                autoFit.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Native,
                autoFit.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                autoFit.PptxToBinary);
            Assert.Contains("normal autofit", autoFit.Note);
        }

        [Fact]
        public void CapabilityContract_ReportsNativeClassicInteractions() {
            foreach (LegacyPptFeature feature in new[] {
                         LegacyPptFeature.Hyperlinks,
                         LegacyPptFeature.Actions
                     }) {
                LegacyPptCapability capability =
                    LegacyPptCapabilityCatalog.Get(feature);
                Assert.Equal(LegacyPptCapabilityState.Native,
                    capability.ImportToEditableModel);
                Assert.Equal(LegacyPptCapabilityState.Native,
                    capability.NewBinaryWrite);
                Assert.Equal(LegacyPptCapabilityState.Preserved,
                    capability.BinaryRoundTrip);
                Assert.Equal(LegacyPptCapabilityState.Blocked,
                    capability.PptxToBinary);
                Assert.Contains("loss-blocked", capability.Note);
            }
        }

        [Fact]
        public void CapabilityContract_ReportsNativeLayoutHintsAndMaterializedConversion() {
            LegacyPptCapability layouts = LegacyPptCapabilityCatalog.Get(LegacyPptFeature.Layouts);

            Assert.Equal(LegacyPptCapabilityState.Native, layouts.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native, layouts.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved, layouts.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Converted, layouts.PptxToBinary);
            Assert.Contains("master-shape visibility maps to schema-valid root attributes and native inheritance flags",
                layouts.Note);
            Assert.Contains("newly added supported layout shapes and placeholders append native OfficeArt shape containers",
                layouts.Note);
            Assert.Contains("explicitly loss-blocked", layouts.Note);

            LegacyPptCapability placeholders = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.Placeholders);
            Assert.Equal(LegacyPptCapabilityState.Native, placeholders.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native, placeholders.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved, placeholders.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native, placeholders.PptxToBinary);
            Assert.Contains("addition, and removal edits on imported slide and main-, title-, notes-, or handout-master placeholders",
                placeholders.Note);
        }

        [Fact]
        public void CapabilityContract_ReportsCompleteDrawingMlThemeConversion() {
            LegacyPptCapability themes = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.Themes);

            Assert.Equal(LegacyPptCapabilityState.Native,
                themes.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                themes.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                themes.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native,
                themes.PptxToBinary);
            Assert.Contains("all twelve colors", themes.Note);
            Assert.Contains("Imported main-, title-, notes-, and handout-master theme edits plus edits to existing imported slide and notes-page theme overrides",
                themes.Note);
            Assert.Contains("Ordinary imported layout theme overrides materialize into every affected slide",
                themes.Note);
        }

        [Fact]
        public void CapabilityContract_ReportsOlePropertySetParity() {
            LegacyPptCapability builtIn = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.BuiltInProperties);
            LegacyPptCapability custom = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.CustomProperties);

            Assert.Equal(LegacyPptCapabilityState.Native,
                builtIn.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                builtIn.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                builtIn.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native,
                builtIn.PptxToBinary);
            Assert.Equal(LegacyPptCapabilityState.Native,
                custom.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                custom.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                custom.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native,
                custom.PptxToBinary);
            Assert.Contains("loss-blocked", builtIn.Note);
            Assert.Contains("byte-preserved", custom.Note);
        }

        [Fact]
        public void CapabilityContract_ReportsBinaryEncryptionParity() {
            LegacyPptCapability encryption = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.Encryption);

            Assert.Equal(LegacyPptCapabilityState.Native,
                encryption.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                encryption.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Native,
                encryption.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native,
                encryption.PptxToBinary);
            Assert.Contains("EncryptedSummary", encryption.Note);
            Assert.Contains("40 through 128 bits", encryption.Note);
            Assert.Contains("legacy interoperability", encryption.Note);
        }

        [Fact]
        public void CapabilityContract_ReportsPreservingMainMasterShapeEdits() {
            LegacyPptCapability masters = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.Masters);

            Assert.Equal(LegacyPptCapabilityState.Preserved,
                masters.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native,
                masters.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                masters.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                masters.PptxToBinary);
            Assert.Contains("Imported main-, title-, notes-, and handout-master position, size, and structurally plain text edits",
                masters.Note);
            Assert.Contains("Binary title masters map to exact named layout parts",
                masters.Note);
            Assert.Contains("parent-master-shape visibility edits",
                masters.Note);
        }

        [Fact]
        public void CapabilityContract_ReportsPreservingMasterBackgroundEdits() {
            LegacyPptCapability backgrounds = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.Backgrounds);

            Assert.Equal(LegacyPptCapabilityState.Native,
                backgrounds.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                backgrounds.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                backgrounds.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                backgrounds.PptxToBinary);
            Assert.Contains("Supported imported slide, notes-page, main-, title-, notes-, and handout-master background edits",
                backgrounds.Note);
            Assert.Contains("ordinary PPTX-layout background edits materialize into every affected imported slide",
                backgrounds.Note);
            Assert.Contains("centered-path, and shape-path gradients",
                backgrounds.Note);
            Assert.Contains("linearly interpolated gradient-stop opacity",
                backgrounds.Note);
            Assert.Contains("picture backgrounds", backgrounds.Note);
            Assert.Contains("deduplicate identical payloads",
                backgrounds.Note);
            Assert.Contains("balance existing BLIP reference counts",
                backgrounds.Note);
        }

        [Fact]
        public void CapabilityContract_FinalizesRichNotesAndUnknownPreservation() {
            LegacyPptCapability richNotes = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.RichNotes);
            Assert.Equal(LegacyPptCapabilityState.Native,
                richNotes.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                richNotes.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                richNotes.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                richNotes.PptxToBinary);

            LegacyPptCapability unknown = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.UnknownRecordsAndStreams);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                unknown.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                unknown.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                unknown.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                unknown.PptxToBinary);
            Assert.Contains("explicitly loss-blocked", unknown.Note);
        }

        [Fact]
        public void CapabilityContract_ReportsHeaderFooterConversionAndPreservingEdits() {
            LegacyPptCapability capability = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.HeadersAndFooters);

            Assert.Equal(LegacyPptCapabilityState.Native,
                capability.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                capability.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                capability.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native,
                capability.PptxToBinary);
        }

        [Fact]
        public void CapabilityContract_SeparatesClassicAndModernComments() {
            LegacyPptCapability classic = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.Comments);
            Assert.Equal(LegacyPptCapabilityState.Native,
                classic.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                classic.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Native,
                classic.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native,
                classic.PptxToBinary);

            LegacyPptCapability modern = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.ModernComments);
            Assert.Equal(LegacyPptRepresentability.NotRepresentable,
                modern.Representability);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                modern.PptxToBinary);
        }

        [Fact]
        public void CapabilityContract_ReportsNativeClassicAndBlockedModernTransitions() {
            LegacyPptCapability transitions = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.Transitions);

            Assert.Equal(LegacyPptCapabilityState.Native,
                transitions.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                transitions.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Native,
                transitions.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                transitions.PptxToBinary);

            LegacyPptCapability sounds = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.TransitionAndActionSounds);
            Assert.Equal(LegacyPptCapabilityState.Native,
                sounds.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                sounds.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Native,
                sounds.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                sounds.PptxToBinary);
        }

        [Fact]
        public void CapabilityContract_ReportsNativeClassicAndBlockedModernAnimations() {
            LegacyPptCapability animations = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.Animations);

            Assert.Equal(LegacyPptCapabilityState.Native,
                animations.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                animations.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Native,
                animations.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                animations.PptxToBinary);
            Assert.Contains("visibility-set scaffold", animations.Note,
                StringComparison.Ordinal);
            Assert.Contains("group-child animation edits", animations.Note,
                StringComparison.Ordinal);
        }
    }
}
