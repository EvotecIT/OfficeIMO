using OfficeIMO.Tool.Commands.Markup;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public class MarkupEmitOutputTests {
    [Fact]
    public async Task EmittedTextRefusesAFileCreatedAfterPreflightUnlessForced() {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO.Tool.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "emitted.cs");
        byte[] original = Enumerable.Repeat((byte)0xA5, 256).ToArray();
        const string emitted = "// Café 📝\n";

        try {
            Assert.False(File.Exists(path));
            await File.WriteAllBytesAsync(path, original);

            await Assert.ThrowsAsync<IOException>(() => MarkupCommand.WriteEmittedTextAsync(path, emitted, force: false));
            Assert.Equal(original, await File.ReadAllBytesAsync(path));

            await MarkupCommand.WriteEmittedTextAsync(path, emitted, force: true);
            Assert.Equal(Encoding.UTF8.GetBytes(emitted), await File.ReadAllBytesAsync(path));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }
}
