using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        [Fact]
        public void LegacySignatures_AreInspectedAndExactNoOpSaveIsAllowed() {
            byte[] sourceBytes = CreateLegacySignatureFixture();

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(input);
            PowerPointSignatureReport report = presentation.InspectSignatures();

            Assert.True(report.HasSignatureMetadata);
            Assert.True(report.HasLegacyBinarySignatureStream);
            Assert.True(report.HasLegacyXmlSignatureStorage);
            Assert.False(report.HasOriginPart);
            Assert.Equal(0, report.XmlSignaturePartCount);
            Assert.Contains("\"hasLegacyBinarySignatureStream\":true", report.ToJson());
            Assert.Equal(sourceBytes, presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void LegacySignatures_BlockEditedSaveByDefaultAndCanBePreservedExplicitly() {
            byte[] sourceBytes = CreateLegacySignatureFixture();

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(input);
            Assert.Single(presentation.Slides[0].TextBoxes,
                textBox => textBox.Text == "Signed deck").Text = "Edited signed deck";

            PowerPointSignedPresentationMutationException blocked =
                Assert.Throws<PowerPointSignedPresentationMutationException>(() =>
                    presentation.ToBytes(PowerPointFileFormat.Ppt));
            Assert.Equal(PowerPointSignatureMutationAction.Blocked, blocked.Report.Action);
            Assert.True(blocked.Report.HasLegacyBinarySignatureStream);
            Assert.True(blocked.Report.HasLegacyXmlSignatureStorage);

            presentation.SignatureMutationPolicy =
                PowerPointSignatureMutationPolicy.PreserveSignatureMarkup;
            byte[] preservedBytes = presentation.ToBytes(PowerPointFileFormat.Ppt);

            LegacyPptPresentation preserved = LegacyPptPresentation.Load(preservedBytes);
            IReadOnlyDictionary<string, byte[]> streams = preserved.Package.CopyCompoundStreams();
            Assert.Equal("binary-signature"u8.ToArray(), streams["_signatures"]);
            Assert.Equal("xml-signature"u8.ToArray(), streams["_xmlsignatures/sig1"]);
            Assert.Equal(PowerPointSignatureMutationAction.Preserved,
                presentation.LastSignatureReport!.Action);
        }

        [Fact]
        public void LegacySignatures_CanBeRemovedExplicitlyFromEditedSave() {
            byte[] sourceBytes = CreateLegacySignatureFixture();

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(input);
            Assert.Single(presentation.Slides[0].TextBoxes,
                textBox => textBox.Text == "Signed deck").Text = "Unsigned edited deck";
            presentation.SignatureMutationPolicy =
                PowerPointSignatureMutationPolicy.RemoveInvalidatedSignatures;

            byte[] unsignedBytes = presentation.ToBytes(PowerPointFileFormat.Ppt);

            LegacyPptPresentation unsigned = LegacyPptPresentation.Load(unsignedBytes);
            Assert.DoesNotContain(unsigned.Package.CopyCompoundStreams().Keys, path =>
                path.Equals("_signatures", StringComparison.OrdinalIgnoreCase)
                || path.Equals("_xmlsignatures", StringComparison.OrdinalIgnoreCase)
                || path.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase));
            Assert.Equal(PowerPointSignatureMutationAction.Removed,
                presentation.LastSignatureReport!.Action);
        }

        private static byte[] CreateLegacySignatureFixture() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Signed deck");
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(sourceBytes);
            return legacy.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) {
                    ["_signatures"] = "binary-signature"u8.ToArray(),
                    ["_xmlsignatures/sig1"] = "xml-signature"u8.ToArray()
                });
        }
    }
}
