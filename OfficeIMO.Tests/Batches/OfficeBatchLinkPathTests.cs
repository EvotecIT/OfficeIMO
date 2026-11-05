namespace OfficeIMO.Tests;

public sealed class OfficeBatchLinkPathTests {
    [Fact]
    public void BatchPathsAcceptSystemRootLinksButRejectLinksInsideTheTree() {
        string root = Directory.CreateTempSubdirectory("officeimo-batch-links-").FullName;
        try {
            string real = Directory.CreateDirectory(Path.Combine(root, "real")).FullName;
            // The default temp location sits below /var on macOS, a link owned by the system.
            OfficeBatchExecutor.EnsureNoLinks(Path.Combine(real, "source.md"));
            if (!OperatingSystem.IsWindows())
                foreach (var link in new DirectoryInfo("/").EnumerateDirectories().Where(entry => entry.LinkTarget != null))
                    OfficeBatchExecutor.EnsureNoLinks(Path.Combine(link.FullName, "officeimo-missing-child"));

            string linked = Path.Combine(root, "linked");
            Directory.CreateSymbolicLink(linked, real);
            Assert.Throws<InvalidDataException>(() => OfficeBatchExecutor.EnsureNoLinks(Path.Combine(linked, "source.md")));
            Assert.Throws<InvalidDataException>(() => OfficeBatchExecutor.EnsureNoLinks(linked));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }
}
