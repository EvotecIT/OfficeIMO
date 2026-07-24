using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioAllSeverityBatch14SecurityTests {
    [Fact]
    public void ShowcaseSummaryRejectsArtifactsOutsideItsRootBeforeHashing() {
        string parent = Path.Combine(Path.GetTempPath(), "officeimo-showcase-boundary-" + Guid.NewGuid().ToString("N"));
        string root = Path.Combine(parent, "root");
        string outside = Path.Combine(parent, "outside.vsdx");
        Directory.CreateDirectory(root);
        File.WriteAllText(outside, "outside");
        try {
            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                VisioShowcaseSummary.Create(root, new[] { outside }));

            Assert.Contains("inside the showcase root", exception.Message, StringComparison.Ordinal);
        } finally {
            Directory.Delete(parent, true);
        }
    }
}
