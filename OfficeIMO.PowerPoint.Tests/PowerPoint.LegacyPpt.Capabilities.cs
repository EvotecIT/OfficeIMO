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
        public void CapabilityContract_SerializesDeterministicallyAndReportsRealGaps() {
            string first = LegacyPptCapabilityCatalog.ToJson();
            string second = LegacyPptCapabilityCatalog.ToJson();

            Assert.Equal(first, second);
            Assert.Contains("\"schemaVersion\":1", first);
            Assert.Contains("\"feature\":\"UnknownRecordsAndStreams\"", first);
            Assert.True(LegacyPptCapabilityCatalog.HasRemainingParityWork);
            Assert.Contains(LegacyPptCapabilityCatalog.RemainingParityWork,
                row => row.Feature == LegacyPptFeature.Masters);
            Assert.Contains(LegacyPptCapabilityCatalog.RemainingParityWork,
                row => row.Feature == LegacyPptFeature.Layouts);
            Assert.Contains("| Preservation | UnknownRecordsAndStreams |",
                LegacyPptCapabilityCatalog.ToMarkdown());
        }

        [Fact]
        public void WritePreflight_FindingsReferenceNonNativeCapabilityRows() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.Hidden = true;
            slide.AddRectangle(100000, 100000, 1000000, 500000).Fill("FF0000");

            LegacyPptWritePreflightReport report = presentation.AnalyzeLegacyPptWrite();

            Assert.NotEmpty(report.Findings);
            foreach (LegacyPptWriteFinding finding in report.Findings) {
                LegacyPptCapability capability = LegacyPptCapabilityCatalog.Get(finding.Feature);
                Assert.NotEqual(LegacyPptCapabilityState.Native, capability.PptxToBinary);
            }
        }

        [Fact]
        public void CapabilityContract_ReportsImplementedRasterImportAndPreservingRoundTrip() {
            LegacyPptCapability raster = LegacyPptCapabilityCatalog.Get(LegacyPptFeature.RasterPictures);

            Assert.Equal(LegacyPptCapabilityState.Native, raster.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Preserved, raster.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Planned, raster.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Planned, raster.PptxToBinary);
        }

        [Fact]
        public void CapabilityContract_ReportsAccessibilityImportAndPreservingRoundTrip() {
            LegacyPptCapability accessibility = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.AccessibilityMetadata);

            Assert.Equal(LegacyPptCapabilityState.Native, accessibility.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Preserved, accessibility.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Planned, accessibility.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Planned, accessibility.PptxToBinary);
        }

        [Fact]
        public void CapabilityContract_ReportsLayoutSubsetWithoutOverstatingCustomLayoutWriting() {
            LegacyPptCapability layouts = LegacyPptCapabilityCatalog.Get(LegacyPptFeature.Layouts);

            Assert.Equal(LegacyPptCapabilityState.Native, layouts.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Planned, layouts.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved, layouts.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Planned, layouts.PptxToBinary);
            Assert.Contains("master-shape visibility maps to schema-valid root attributes and native inheritance flags",
                layouts.Note);

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
        public void CapabilityContract_ReportsPreservingMainMasterShapeEdits() {
            LegacyPptCapability masters = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.Masters);

            Assert.Equal(LegacyPptCapabilityState.Planned,
                masters.BinaryRoundTrip);
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

            Assert.Contains("Supported imported slide, notes-page, main-, title-, notes-, and handout-master background edits",
                backgrounds.Note);
            Assert.Contains("ordinary PPTX-layout background edits materialize into every affected imported slide",
                backgrounds.Note);
            Assert.Contains("centered-path, and shape-path gradients",
                backgrounds.Note);
            Assert.Contains("linearly interpolated gradient-stop opacity",
                backgrounds.Note);
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
