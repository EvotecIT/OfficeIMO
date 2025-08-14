using System;
using System.Runtime.InteropServices;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public class FontResolverTests {
    [Fact]
    public void Resolve_GenericFonts() {
        string resolved = FontResolver.Resolve("monospace")!;

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            Assert.Contains(resolved, new[] { "Consolas", "Calibri" });
        } else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
            Assert.Contains(resolved, new[] { "DejaVu Sans Mono", "DejaVu Sans" });
        } else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            Assert.Contains(resolved, new[] { "Menlo", "Helvetica" });
        }
    }

    [Fact]
    public void Resolve_FallbackFonts() {
        string resolved = FontResolver.Resolve("DefinitelyMissingFont")!;

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            Assert.Equal("Calibri", resolved);
        } else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
            Assert.Equal("DejaVu Sans", resolved);
        } else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            Assert.Equal("Helvetica", resolved);
        }
    }

    [Theory]
    [InlineData("cursive")]
    [InlineData("fantasy")]
    public void Resolve_ExtendedGenericFonts(string generic) {
        string resolved = FontResolver.Resolve(generic)!;
        Assert.False(string.Equals(resolved, generic, StringComparison.OrdinalIgnoreCase));
    }
}

