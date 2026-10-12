using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class OfficeCloudFontDiscoveryTests {
    [Fact]
    public void CachedFamilyDiscoveryAcceptsNumericFontFilenamesAndKeepsFamiliesSeparate() {
        string cache = Path.Combine(Path.GetTempPath(), "OfficeIMO.CloudFontDiscovery", Guid.NewGuid().ToString("N"));
        try {
            Assert.Empty(OfficeCloudFontFiles.EnumerateRoots(cache));
            string cloud = Path.Combine(cache, "4", "CloudFonts");
            string family = Path.Combine(cloud, "Example Sans");
            string other = Path.Combine(cloud, "Other Family");
            Directory.CreateDirectory(family);
            Directory.CreateDirectory(other);
            string numeric = Path.Combine(family, "123456789.ttf");
            // Discovery returns candidates. The existing font loader validates
            // their programs and declared family before selecting a face.
            File.WriteAllBytes(numeric, new byte[] { 0, 1 });
            File.WriteAllBytes(Path.Combine(family, "metadata.json"), new byte[] { 0, 1 });
            File.WriteAllBytes(Path.Combine(other, "987654321.ttf"), new byte[] { 0, 1 });

            Assert.Equal(new[] { cloud }, OfficeCloudFontFiles.EnumerateRoots(cache).ToArray());
            Assert.Equal(new[] { numeric }, OfficeCloudFontFiles.EnumerateFamilyFiles(cloud, "examplesans").ToArray());
            Assert.Empty(OfficeCloudFontFiles.EnumerateFamilyFiles(cloud, "missingfamily"));
        } finally {
            if (Directory.Exists(cache)) Directory.Delete(cache, recursive: true);
        }
    }
}
