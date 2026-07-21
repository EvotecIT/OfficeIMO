using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointOfficeInteroperabilityCorpusTests {
    [Fact]
    [Trait("Category", "OfficeInteroperability")]
    public void CorpusManifest_TracksAnalyzesAndConvertsEveryLegacyPresentation() {
        OfficeInteroperabilityCorpusManifest manifest = OfficeInteroperabilityCorpusManifestLoader.Load();
        IReadOnlyList<string> errors = OfficeInteroperabilityCorpusManifestLoader.Validate(manifest);
        Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors));

        OfficeInteroperabilityCorpusCollection collection = Assert.Single(
            manifest.Collections,
            item => item.Format == "ppt");
        Assert.Equal("PowerPoint.Ppt", collection.FormatId);
        Assert.Contains("legacy-to-modern", collection.Directions);
        Assert.Contains("visual", collection.Oracles);
        Assert.Equal(12, collection.Artifacts.Count);

        string outputDirectory = Path.Combine(
            Path.GetTempPath(),
            "OfficeIMO-PowerPoint-Corpus-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outputDirectory);
        try {
            foreach (OfficeInteroperabilityCorpusArtifact artifact in collection.Artifacts) {
                string source = OfficeInteroperabilityCorpusManifestLoader.ResolveArtifactPath(collection, artifact);
                string destination = Path.Combine(
                    outputDirectory,
                    Path.GetFileNameWithoutExtension(artifact.File) + ".pptx");

                using PowerPointPresentation imported = PowerPointPresentation.Load(source);
                Assert.Equal("PowerPoint.Ppt", imported.SourceFormatDescriptor.Id);
                int expectedSlideCount = imported.Slides.Count;
                Assert.True(expectedSlideCount > 0, artifact.File);

                var options = new PowerPointPresentationConversionOptions {
                    CompatibilityMode = OfficeCompatibilityMode.BestEffort,
                    LossPolicy = PowerPointConversionLossPolicy.Allow
                };
                PowerPointPresentationConversionReport preflight =
                    PowerPointPresentation.AnalyzeConversion(source, destination, options);
                Assert.Equal("PowerPoint.Ppt", preflight.SourceFormatDescriptor.Id);
                Assert.Equal("PowerPoint.Pptx", preflight.DestinationFormatDescriptor.Id);
                Assert.All(preflight.Compatibility.Findings, finding => {
                    Assert.False(string.IsNullOrWhiteSpace(finding.Code));
                    Assert.False(string.IsNullOrWhiteSpace(finding.Category));
                });
                Assert.False(File.Exists(destination));

                PowerPointPresentationConversionResult conversion =
                    PowerPointPresentation.Convert(source, destination, options);
                Assert.Equal(destination, conversion.RequireValue());
                Assert.Equal(
                    preflight.Compatibility.Findings.Select(finding => finding.Code),
                    conversion.Report.Compatibility.Findings.Select(finding => finding.Code));

                using PowerPointPresentation reopened = PowerPointPresentation.Load(destination);
                Assert.Equal(expectedSlideCount, reopened.Slides.Count);
            }
        } finally {
            try {
                Directory.Delete(outputDirectory, recursive: true);
            } catch {
            }
        }
    }
}
