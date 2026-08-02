using System;
using System.IO;
using System.Linq;
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
    }
}
