using OfficeIMO.Core;
using Xunit;

namespace OfficeIMO.Tests;

public class OfficeCoreContractTests {
    [Fact]
    public void FileLauncherRejectsMissingFilesBeforeStartingAnApplication() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        FileNotFoundException exception = Assert.Throws<FileNotFoundException>(() => OfficeFileLauncher.Open(path));
        Assert.Equal(Path.GetFullPath(path), exception.FileName);
    }
}
