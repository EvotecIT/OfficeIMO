using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        [Fact]
        public void PublicMacroApi_AddsReplacesRemovesAndBoundsVbaProject() {
            byte[] original = CreateVbaTestProject("OriginalModule",
                "Sub OriginalMacro()\nEnd Sub");
            byte[] replacement = CreateVbaTestProject("ReplacementModule",
                "Sub ReplacementMacro()\nEnd Sub");

            using var presentation = PowerPointPresentation.Create();
            Assert.False(presentation.HasVbaProject);
            Assert.Null(presentation.GetVbaProjectBytes());
            Assert.False(presentation.RemoveVbaProject());

            presentation.SetVbaProject(original);
            Assert.True(presentation.HasVbaProject);
            Assert.Equal(original, presentation.GetVbaProjectBytes());
            Assert.Contains(presentation.InspectFeatures().EditableFeatures,
                feature => feature.Name == "VBA macros" && feature.Count == 1);

            using var replacementStream = new MemoryStream(replacement,
                writable: false);
            replacementStream.Position = Math.Min(7, replacement.Length);
            presentation.SetVbaProject(replacementStream);
            Assert.Equal(Math.Min(7, replacement.Length), replacementStream.Position);
            Assert.Equal(replacement, presentation.GetVbaProjectBytes());

            Assert.Throws<InvalidDataException>(() =>
                presentation.GetVbaProjectBytes(replacement.Length - 1));
            Assert.Throws<InvalidDataException>(() =>
                presentation.SetVbaProject(replacement, replacement.Length - 1));
            Assert.Throws<InvalidDataException>(() =>
                presentation.SetVbaProject(new byte[] { 1, 2, 3, 4 }));

            Assert.True(presentation.RemoveVbaProject());
            Assert.False(presentation.HasVbaProject);
            Assert.Null(presentation.GetVbaProjectBytes());
        }

        [Fact]
        public void PublicMacroApi_RoundTripsMacroEnabledAndLegacyDestinations() {
            string pptm = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-MacroApi-" + Guid.NewGuid().ToString("N") + ".pptm");
            byte[] project = CreateVbaTestProject("RoundTripModule",
                "Sub RoundTripMacro()\nEnd Sub");
            try {
                using (PowerPointPresentation presentation =
                       PowerPointPresentation.Create(pptm)) {
                    presentation.AddSlide().AddTitle("Macro-enabled deck");
                    presentation.SetVbaProject(project);
                    presentation.Save();
                }

                using PowerPointPresentation reopened =
                    PowerPointPresentation.Load(pptm);
                Assert.Equal(project, reopened.GetVbaProjectBytes());

                byte[] binary = reopened.ToBytes(PowerPointFileFormat.Ppt);
                using PowerPointPresentation projected =
                    PowerPointPresentation.Load(new MemoryStream(binary));
                Assert.True(projected.HasVbaProject,
                    string.Join(Environment.NewLine, projected.LegacyPptImportDiagnostics
                        .Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message)));
                Assert.Equal(project, projected.GetVbaProjectBytes());
            } finally {
                if (File.Exists(pptm)) File.Delete(pptm);
            }
        }

        [Fact]
        public void PublicMacroApi_RetainsFreshProjectsInStreamOutputs() {
            byte[] project = CreateVbaTestProject("StreamModule",
                "Sub StreamMacro()\nEnd Sub");
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.AddSlide().AddTitle("Macro stream output");
            presentation.SetVbaProject(project);

            AssertMacroEnabledBytes(presentation.ToBytes(), project);
            AssertMacroEnabledBytes(
                presentation.ToBytes(PowerPointFileFormat.Pptm), project);
            using var inferredStream = new MemoryStream();
            presentation.Save(inferredStream);
            AssertMacroEnabledBytes(inferredStream.ToArray(), project);
            using var stream = new MemoryStream();
            presentation.Save(stream, PowerPointFileFormat.Pptm);
            AssertMacroEnabledBytes(stream.ToArray(), project);
            byte[] encrypted = presentation.ToEncryptedBytes(
                "macro-stream-pass", PowerPointFileFormat.Pptm);
            AssertEncryptedMacroEnabledBytes(encrypted,
                "macro-stream-pass", project);
            using var encryptedStream = new MemoryStream();
            presentation.SaveEncrypted(encryptedStream,
                "macro-save-pass", PowerPointFileFormat.Pptm);
            AssertEncryptedMacroEnabledBytes(encryptedStream.ToArray(),
                "macro-save-pass", project);

            using PowerPointPresentation pptx = PowerPointPresentation.Load(
                new MemoryStream(presentation.ToBytes(
                    PowerPointFileFormat.Pptx)));
            Assert.False(pptx.HasVbaProject);
            Assert.Equal(PresentationDocumentType.Presentation,
                pptx.OpenXmlDocument.DocumentType);
        }

        [Fact]
        public void PowerPointFileFormat_PreservesLegacyOrdinalsAndAddsPptm() {
            Assert.Equal(0, (int)PowerPointFileFormat.Pptx);
            Assert.Equal(1, (int)PowerPointFileFormat.Ppt);
            Assert.Equal(2, (int)PowerPointFileFormat.Pot);
            Assert.Equal(3, (int)PowerPointFileFormat.Pps);
            Assert.Equal(4, (int)PowerPointFileFormat.Pptm);
        }

        [Fact]
        public void PublicMacroApi_RejectsCorruptVbaHeadersAndDirectories() {
            byte[] badHeader = CreateVbaTestProject("BadHeader", "Sub Main(): End Sub",
                corruptProjectHeader: true);
            byte[] badDirectory = CreateVbaTestProject("BadDirectory", "Sub Main(): End Sub",
                corruptDirectory: true);
            byte[] badDirectoryRecords = CreateVbaTestProject(
                "BadDirectoryRecords", "Sub Main(): End Sub",
                corruptDirectoryRecords: true);
            byte[] missingModule = CreateVbaTestProject(
                "MissingModule", "Sub Main(): End Sub",
                omitModuleStream: true);
            using PowerPointPresentation presentation = PowerPointPresentation.Create();

            InvalidDataException headerError = Assert.Throws<InvalidDataException>(() =>
                presentation.SetVbaProject(badHeader));
            InvalidDataException directoryError = Assert.Throws<InvalidDataException>(() =>
                presentation.SetVbaProject(badDirectory));
            InvalidDataException recordError = Assert.Throws<InvalidDataException>(() =>
                presentation.SetVbaProject(badDirectoryRecords));
            InvalidDataException moduleError = Assert.Throws<InvalidDataException>(() =>
                presentation.SetVbaProject(missingModule));

            Assert.Contains("_VBA_PROJECT", headerError.Message,
                StringComparison.Ordinal);
            Assert.Contains("dir", directoryError.Message,
                StringComparison.Ordinal);
            Assert.Contains("dir", recordError.Message,
                StringComparison.Ordinal);
            Assert.Contains("dir", moduleError.Message,
                StringComparison.Ordinal);
            Assert.False(presentation.HasVbaProject);
        }

        [Fact]
        public void FeatureReport_PreservesVbaProjectsWithRelatedParts() {
            byte[] project = CreateVbaTestProject("RelatedPartModule",
                "Sub Main(): End Sub");
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.AddSlide().AddTitle("Related VBA parts");
            presentation.SetVbaProject(project);
            VbaDataPart dataPart = presentation.OpenXmlDocument
                .PresentationPart!.VbaProjectPart!
                .AddNewPart<VbaDataPart>();
            using (var data = new MemoryStream(new byte[] { 1, 2, 3, 4 },
                       writable: false)) {
                dataPart.FeedData(data);
            }

            PowerPointFeatureReport report = presentation.InspectFeatures();
            PowerPointFeatureFinding macros = Assert.Single(
                report.FindFeatures("VBA macros"));

            Assert.Equal(PowerPointFeatureSupportLevel.Preserved,
                macros.SupportLevel);
            Assert.Throws<InvalidOperationException>(() =>
                report.EnsureNoAdvancedFeatures());
        }

        private static void AssertMacroEnabledBytes(byte[] bytes,
            byte[] expectedProject) {
            using PowerPointPresentation loaded = PowerPointPresentation.Load(
                new MemoryStream(bytes));
            Assert.Equal(PresentationDocumentType.MacroEnabledPresentation,
                loaded.OpenXmlDocument.DocumentType);
            Assert.Equal(PowerPointFileFormat.Pptm, loaded.SourceFormat);
            Assert.Equal(expectedProject, loaded.GetVbaProjectBytes());
        }

        private static void AssertEncryptedMacroEnabledBytes(byte[] bytes,
            string password, byte[] expectedProject) {
            using PowerPointPresentation loaded =
                PowerPointPresentation.LoadEncrypted(
                    new MemoryStream(bytes), password);
            Assert.Equal(PresentationDocumentType.MacroEnabledPresentation,
                loaded.OpenXmlDocument.DocumentType);
            Assert.Equal(PowerPointFileFormat.Pptm, loaded.SourceFormat);
            Assert.Equal(expectedProject, loaded.GetVbaProjectBytes());
        }
    }
}
