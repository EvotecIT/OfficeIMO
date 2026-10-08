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
            Assert.Null(fonts.CreateSubstitutionDiagnostic(text, family));

            // Installed availability is platform-dependent; the direct generic-family
            // contract above also applies when no corresponding font is installed.
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
