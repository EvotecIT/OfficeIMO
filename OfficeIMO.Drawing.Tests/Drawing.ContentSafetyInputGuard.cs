using OfficeIMO.ContentSafety;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ContentSafetyInputGuardTests {
    [Fact]
    public void DefaultLiteralStillBindsToTheExistingZipInspectionParameter() {
        string path = Path.GetTempFileName();
        try {
            File.WriteAllBytes(path, new byte[] { 1, 2, 3 });

            byte[] bytes = OfficeContentSafetyInputGuard.ReadAllBytes(
                path,
                new OfficeContentSafetyOptions(),
                default);

            Assert.Equal(new byte[] { 1, 2, 3 }, bytes);
        } finally {
            File.Delete(path);
        }
    }
}
