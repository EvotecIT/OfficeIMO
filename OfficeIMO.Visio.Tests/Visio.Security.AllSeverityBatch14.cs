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

#if NET8_0_OR_GREATER
    [Fact]
    public void ShowcaseSummaryRejectsSymlinkedArtifactsAndParentDirectories() {
        string parent = Path.Combine(Path.GetTempPath(), "officeimo-showcase-symlink-" + Guid.NewGuid().ToString("N"));
        string root = Path.Combine(parent, "root");
        string outsideDirectory = Path.Combine(parent, "outside");
        Directory.CreateDirectory(root);
        Directory.CreateDirectory(outsideDirectory);
        string outside = Path.Combine(outsideDirectory, "outside.vsdx");
        File.WriteAllText(outside, "outside");
        string fileLink = Path.Combine(root, "file-link.vsdx");
        string directoryLink = Path.Combine(root, "directory-link");
        try {
            File.CreateSymbolicLink(fileLink, outside);
            Directory.CreateSymbolicLink(directoryLink, outsideDirectory);

            Assert.Throws<ArgumentException>(() =>
                VisioShowcaseSummary.Create(root, new[] { fileLink }));
            Assert.Throws<ArgumentException>(() =>
                VisioShowcaseSummary.Create(root, new[] { Path.Combine(directoryLink, "outside.vsdx") }));
        } finally {
            Directory.Delete(parent, true);
        }
    }
#endif
}
