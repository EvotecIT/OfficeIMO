using System;
using System.Runtime.InteropServices;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public class FontResolverTests {
    [Theory]
    [InlineData("monospace", "Consolas")]
    [InlineData(" Monospace ", "Consolas")]
    [InlineData("serif", "Times New Roman")]
    [InlineData("sans-serif", "Calibri")]
    public void Resolve_GenericFontsToPortableDocumentFamilies(string generic, string expected) {
        // The written family must not depend on the machine that generates the document.
        Assert.Equal(expected, WordFontResolver.Resolve(generic));
    }

    [Fact]
    public void Resolve_FallbackFonts() {
        string resolved = WordFontResolver.Resolve("DefinitelyMissingFont")!;

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
        string resolved = WordFontResolver.Resolve(generic)!;
        Assert.False(string.Equals(resolved, generic, StringComparison.OrdinalIgnoreCase));
    }
}

