using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class DrawingTests {
        [Theory]
        [InlineData("system-ui")]
        [InlineData("-apple-system")]
        [InlineData("BlinkMacSystemFont")]
        [InlineData("math")]
        public void GenericInstalledAliasesDoNotFailStrictExportAsNamedSubstitutions(string family) {
            const string text = "OfficeIMO";
            var fonts = new OfficeFontFaceCollection();
            OfficeImageExportDiagnostic? direct = fonts.CreateSubstitutionDiagnostic(text, family);
            OfficeTrueTypeFont? platform = OfficeTrueTypeFont.TryLoadFontFamilyForText(
                family, OfficeFontFaceDescriptor.Regular, text, out _);
            if (platform != null) Assert.Null(direct);
            else {
                Assert.NotNull(direct);
                Assert.Throws<OfficeImageExportPolicyException>(() =>
                    new OfficeImageExportPolicy { RequireNoLoss = true }.EnsureAccepted(new[] { direct! }));
            }

            // Installed CFF math can be registered even where the direct TrueType
            // fallback is unavailable. Both paths preserve their actual diagnostics.
            if (!fonts.TryAddInstalledFamily(family, OfficeFontFaceDescriptor.Regular,
                    128 * 1024 * 1024, System.Threading.CancellationToken.None, out _, out _)) return;

            Assert.Null(fonts.CreateSubstitutionDiagnostic(text, family));
            Assert.Null(fonts.Clone().CreateSubstitutionDiagnostic(text, family));
            Assert.Null(new OfficeFontFaceCollection().AddRange(fonts)
                .CreateSubstitutionDiagnostic(text, family));
            var drawing = new OfficeDrawing(160, 32);
            drawing.Fonts.AddRange(fonts);
            drawing.AddText(text, 0, 0, 160, 32, new OfficeFontInfo(family, 16), OfficeColor.Black);
            var diagnostics = new System.Collections.Generic.List<OfficeImageExportDiagnostic>();
            drawing.AppendFontDiagnostics(diagnostics);
            new OfficeImageExportPolicy { RequireNoLoss = true }.EnsureAccepted(diagnostics);
            Assert.Empty(diagnostics);
        }
    }
}
