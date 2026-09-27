using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSystemUiFontTests {
    [Theory]
    [InlineData("system-ui")]
    [InlineData("-apple-system")]
    [InlineData("BlinkMacSystemFont")]
    public void MacSystemUiNamesSelectSfnsForEmbeddingAndDescriptors(string familyName) {
        const string path = "/System/Library/Fonts/SFNS.ttf";
        if (!File.Exists(path)) return;
        Assert.True(PdfEmbeddedFontFamily.TryFromSystem(familyName, out var family));
        Assert.Equal(File.ReadAllBytes(path), family!.Regular);
        Assert.Equal(familyName, family.FamilyName);
        Assert.True(PdfEmbeddedFontFamily.TryResolveSystemFace(familyName,
            OfficeFontFaceDescriptor.Regular, "Animation", out var selected));
        Assert.Equal(family!.Regular, selected!.Regular);
    }
}
