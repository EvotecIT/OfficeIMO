using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.TestAssets;
using System.Runtime.InteropServices;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointLegacyPptComValidationTests {
    [LegacyPptComFact]
    [Trait("Category", "MicrosoftOfficeInteroperability")]
    public void CorpusSourcesAndConvertedPresentationsRenderInDesktopPowerPointWhenRequested() {
        Assert.True(RuntimeInformation.IsOSPlatform(OSPlatform.Windows), "Legacy PPT COM validation requires Windows.");

        OfficeInteroperabilityCorpusManifest manifest = OfficeInteroperabilityCorpusManifestLoader.Load();
        IReadOnlyList<string> errors = OfficeInteroperabilityCorpusManifestLoader.Validate(manifest);
        Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors));
        OfficeInteroperabilityCorpusCollection collection = Assert.Single(
            manifest.Collections,
            item => item.Format == "ppt");

        string root = Path.Combine(
            Path.GetTempPath(),
            "OfficeIMO-PowerPoint-Desktop-Oracle-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            foreach (OfficeInteroperabilityCorpusArtifact artifact in collection.Artifacts) {
                string source = OfficeInteroperabilityCorpusManifestLoader.ResolveArtifactPath(collection, artifact);
                string fixtureName = Path.GetFileNameWithoutExtension(artifact.File);
                AssertPowerPointDesktopRenders(
                    source,
                    Path.Combine(root, fixtureName, "source"),
                    artifact.File + " source");

                string converted = Path.Combine(root, fixtureName, fixtureName + ".pptx");
                PowerPointPresentation.Convert(
                    source,
                    converted,
                    new PowerPointPresentationConversionOptions {
                        CompatibilityMode = OfficeCompatibilityMode.BestEffort,
                        LossPolicy = PowerPointConversionLossPolicy.Allow
                    });
                AssertPowerPointDesktopRenders(
                    converted,
                    Path.Combine(root, fixtureName, "converted"),
                    artifact.File + " converted PPTX");
            }


            string modernSource = Path.Combine(root, "officeimo-modern-source.pptx");
            string binaryDestination = Path.Combine(root, "officeimo-modern-to-binary.ppt");
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(modernSource)) {
                presentation.AddSlide().AddTextBox("OfficeIMO modern to binary");
                presentation.Save();
            }
            PowerPointPresentation.Convert(modernSource, binaryDestination).RequireNoLoss();
            AssertPowerPointDesktopRenders(
                binaryDestination,
                Path.Combine(root, "officeimo-modern-to-binary-render"),
                "OfficeIMO-generated PPT output");
        } finally {
            try {
                Directory.Delete(root, recursive: true);
            } catch {
            }
        }
    }

    private static void AssertPowerPointDesktopRenders(
        string presentationPath,
        string outputDirectory,
        string label) {
        PowerPointReferenceRenderResult result = PowerPointDesktopReferenceRenderer.TryRender(
            presentationPath,
            outputDirectory,
            enabled: true);
        Assert.Equal(
            PowerPointReferenceRenderStatus.Succeeded,
            result.Status);
        Assert.NotEmpty(result.ImagePaths);
        Assert.All(result.ImagePaths, path =>
            Assert.True(new FileInfo(path).Length > 0, $"{label}: PowerPoint emitted an empty image at '{path}'."));
    }
}
