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
    }
}
