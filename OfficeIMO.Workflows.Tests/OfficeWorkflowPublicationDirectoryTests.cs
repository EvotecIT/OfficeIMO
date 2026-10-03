namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeWorkflowPublicationDirectoryTests {
    [Fact]
    public void Directory_anchored_publication_creates_readable_owner_only_output() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-publication-mode-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        root = OfficeWorkflowPathIdentity.ResolvePhysicalPath(root);
        try {
            string destination = Path.Combine(root, "evidence.json");
            using OfficeWorkflowPublicationDirectory directory = OfficeWorkflowPublicationDirectory.Open(destination);
            directory.WriteNew("staging.tmp", new byte[] { 1, 2, 3 });
            if (!OperatingSystem.IsWindows()) {
                Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite,
                    File.GetUnixFileMode(Path.Combine(root, "staging.tmp")));
            }
            directory.MoveNoReplace("staging.tmp", "evidence.json");
            Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(destination));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }
}
