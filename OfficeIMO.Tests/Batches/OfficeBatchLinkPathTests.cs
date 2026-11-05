namespace OfficeIMO.Tests;

public sealed class OfficeBatchLinkPathTests {
    [Fact]
    public void BatchPathsAcceptSystemRootLinksButRejectLinksInsideTheTree() {
        string root = Directory.CreateTempSubdirectory("officeimo-batch-links-").FullName;
        try {
            string real = Directory.CreateDirectory(Path.Combine(root, "real")).FullName;
            string source = Path.Combine(real, "source.md");
            File.WriteAllText(source, "# Source");
            // The default temp location sits below /var on macOS, a link owned by the system.
            OfficeBatchExecutor.EnsureNoLinks(source);
            using (OfficeBatchExecutor.OpenWithinRoot(source, real, 32)) { }
            // macOS links /tmp and /var to /private; both are ordinary batch locations.
            foreach (string link in new[] { "/tmp", "/var" }.Where(path => !OperatingSystem.IsWindows() && new DirectoryInfo(path).LinkTarget != null))
                OfficeBatchExecutor.EnsureNoLinks(Path.Combine(link, "officeimo-missing-child"));

            string linked = Path.Combine(root, "linked");
            try {
                Directory.CreateSymbolicLink(linked, real);
            } catch (Exception error) when (OperatingSystem.IsWindows() && error is UnauthorizedAccessException or IOException) {
                return; // This Windows account cannot create symbolic links.
            }
            Assert.Throws<InvalidDataException>(() => OfficeBatchExecutor.EnsureNoLinks(Path.Combine(linked, "source.md")));
            Assert.Throws<InvalidDataException>(() => OfficeBatchExecutor.EnsureNoLinks(linked));
            // A root replaced by a link after validation must not redirect reads.
            Assert.Throws<InvalidDataException>(() => OfficeBatchExecutor.OpenWithinRoot(Path.Combine(linked, "source.md"), linked, 32).Dispose());
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }
}
