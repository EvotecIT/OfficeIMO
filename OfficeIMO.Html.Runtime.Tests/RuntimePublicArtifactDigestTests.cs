using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Html.Runtime.Tests;

public sealed class RuntimePublicArtifactDigestTests {
    [Fact]
    public void PublishedPayloadDigestIncludesDependencyBytesAndFileSet() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-public-digest-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            File.WriteAllText(Path.Combine(directory, "renderer.dll"), "entry");
            File.WriteAllText(Path.Combine(directory, "dependency.dll"), "first");
            string original = HtmlPublicArtifactDigest.DirectorySha256(directory);
            File.WriteAllText(Path.Combine(directory, "dependency.dll"), "other");
            string changedDependency = HtmlPublicArtifactDigest.DirectorySha256(directory);
            Assert.NotEqual(original, changedDependency);
            File.WriteAllText(Path.Combine(directory, "unexpected.txt"), "x");
            Assert.NotEqual(changedDependency, HtmlPublicArtifactDigest.DirectorySha256(directory));
        } finally { Directory.Delete(directory, recursive: true); }
    }
}
