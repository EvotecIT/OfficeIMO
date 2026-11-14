namespace OfficeIMO.Workflows.Tests;

public sealed class PrintOutputFileTests {
    [Fact]
    public void AnUnfinishedJobRemovesOnlyTheEmptyFileItCreated() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-print-output-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string created = Path.Combine(directory, "created.pdf");
            File.WriteAllBytes(created, []);
            Assert.True(PrintOutputFile.DiscardEmpty(created, existedBefore: false));
            Assert.False(File.Exists(created));

            string written = Path.Combine(directory, "written.pdf");
            File.WriteAllBytes(written, [0x25]);
            Assert.False(PrintOutputFile.DiscardEmpty(written, existedBefore: false));
            Assert.True(File.Exists(written));

            string existing = Path.Combine(directory, "existing.pdf");
            File.WriteAllBytes(existing, []);
            Assert.False(PrintOutputFile.DiscardEmpty(existing, existedBefore: true));
            Assert.True(File.Exists(existing));

            Assert.False(PrintOutputFile.DiscardEmpty(Path.Combine(directory, "never-created.pdf"), existedBefore: false));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }
}
