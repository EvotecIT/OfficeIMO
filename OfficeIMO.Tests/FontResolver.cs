using System.Runtime.InteropServices;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public class FontResolverTests {
    [Fact]
    public void Resolve_GenericFonts() {
        string resolved = FontResolver.Resolve("monospace")!;

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            Assert.Equal("Consolas", resolved);
        } else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
            Assert.Equal("DejaVu Sans Mono", resolved);
        } else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            Assert.Equal("Menlo", resolved);
        }
    }
}

