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

        [Theory]
        [InlineData("HeiseiMin", "機密")]
        [InlineData("STSong", "秘密")]
        [InlineData("MSung", "秘密")]
        [InlineData("HYSMyeongJo", "비밀")]
        public void InstalledCjkSubstitutionRemainsVisibleToStrictExportPolicies(string family, string text) {
            OfficeTrueTypeFont? reference = OfficeTrueTypeFont.TryLoad(
                "/System/Library/Fonts/Supplemental/Arial Unicode.ttf");
            if (reference == null || !((IOfficeFontProgram)reference).HasGlyphs(text)) return;

            OfficeImageExportDiagnostic? diagnostic = new OfficeFontFaceCollection()
                .CreateSubstitutionDiagnostic(text, family);

            Assert.NotNull(diagnostic);
            Assert.Equal(OfficeImageExportDiagnosticCodes.FontSubstituted, diagnostic!.Code);
            Assert.Equal(OfficeIMO.OfficeConversionLossKind.Approximation, diagnostic.LossKind);
            Assert.Throws<OfficeImageExportPolicyException>(() =>
                new OfficeImageExportPolicy { RequireNoLoss = true }.EnsureAccepted(new[] { diagnostic }));
            Assert.Throws<OfficeImageExportPolicyException>(() =>
                new OfficeImageExportPolicy { FailOnDiagnosticCodes = new[] { diagnostic.Code } }
                    .EnsureAccepted(new[] { diagnostic }));
        }

        [Theory]
        [InlineData("HeiseiMin", "機密")]
        [InlineData("STSong", "秘密")]
        [InlineData("MSung", "秘密")]
        [InlineData("HYSMyeongJo", "비밀")]
        public void InstalledCjkRegularFamilyAndRegistrationUseTheNumericFace(string family, string text) {
            OfficeTrueTypeFont? reference = OfficeTrueTypeFont.TryLoad(
                "/System/Library/Fonts/Supplemental/Arial Unicode.ttf");
            if (reference == null || !((IOfficeFontProgram)reference).HasGlyphs(text)) return;

            OfficeTrueTypeFont? numeric = OfficeTrueTypeFont.TryLoadFontFamilyForText(
                family, OfficeFontFaceDescriptor.Regular, text, out _);
            OfficeTrueTypeFont? legacy = OfficeTrueTypeFont.TryLoadFontFamily(family);
            Assert.NotNull(numeric);
            Assert.NotNull(legacy);
            Assert.Equal(numeric!.FaceDescriptor, legacy!.FaceDescriptor);

            var fonts = new OfficeFontFaceCollection();
            Assert.True(fonts.TryAddInstalledFamily(family, OfficeFontFaceDescriptor.Regular,
                128 * 1024 * 1024, System.Threading.CancellationToken.None, out _, out _));
            Assert.True(fonts.TryResolveFaceForText(text, family, OfficeFontFaceDescriptor.Regular,
                out OfficeFontFace? face));
            Assert.Equal(numeric.FaceDescriptor, face!.Descriptor);
            Assert.NotNull(fonts.CreateSubstitutionDiagnostic(text, family));
            Assert.NotNull(fonts.Clone().CreateSubstitutionDiagnostic(text, family));
            Assert.NotNull(new OfficeFontFaceCollection().AddRange(fonts)
                .CreateSubstitutionDiagnostic(text, family));
        }
    }
}
