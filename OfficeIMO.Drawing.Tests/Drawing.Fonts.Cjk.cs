using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class DrawingTests {
        [Theory]
        [InlineData("HeiseiMin", "機密")]
        [InlineData("STSong", "秘密")]
        [InlineData("MSung", "秘密")]
        [InlineData("HYSMyeongJo", "비밀")]
        public void InstalledCjkSubstitutesCoverTextWhenMacUnicodeReferenceIsAvailable(string family, string text) {
            // This independent installed face is an availability oracle, not a shipped
            // fixture or a requirement on other platforms. Native acceptance records
            // whether the reference is actually present on the validation host.
            OfficeTrueTypeFont? reference = OfficeTrueTypeFont.TryLoad(
                "/System/Library/Fonts/Supplemental/Arial Unicode.ttf");
            if (reference == null || !((IOfficeFontProgram)reference).HasGlyphs(text)) return;

            OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoadFontFamilyForText(
                family, OfficeFontFaceDescriptor.Regular, text, out _);

            Assert.NotNull(font);
            Assert.True(((IOfficeFontProgram)font!).HasGlyphs(text));
            Assert.NotEmpty(font.GetTextContours(text, 0, 0, 18));
        }
    }
}
