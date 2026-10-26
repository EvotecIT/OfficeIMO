using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlDefaultFont_HostRegularDoesNotSuppressPhysicalBoldAndItalicRequests() {
        string? family = new[] { "Arial", "Liberation Sans", "DejaVu Sans" }
            .FirstOrDefault(HasDistinctHostFaces);
        // Installed-font integration is conditional; scoped fixture contracts do
        // not depend on a host supplying three physical styles for one family.
        if (family == null) return;
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<style>body,p{margin:0;font-size:16px}</style>"
            + "<p>REGULAR <strong>BOLD</strong> <em>ITALIC</em></p>");
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, DefaultFontFamily = family,
            ViewportWidth = 360D, ViewportHeight = 100D, Margins = HtmlRenderMargins.All(0D)
        };

        AssertPhysicalFaces(HtmlRenderEngine.Render(source, options));
        // Target preparation must retain the same requested styles without
        // treating its library-only policy as arbitrary document-font access.
        AssertPhysicalFaces(HtmlRenderEngine.Render(source, new HtmlToPdfOptions(options)));

        void AssertPhysicalFaces(HtmlRenderDocument rendered) {
            var text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
                .OfType<HtmlRenderText>().ToArray();
            HtmlRenderText regular = Assert.Single(text, run => run.Text.Contains("REGULAR", StringComparison.Ordinal));
            HtmlRenderText bold = Assert.Single(text, run => run.Text == "BOLD");
            HtmlRenderText italic = Assert.Single(text, run => run.Text == "ITALIC");
            Assert.False(regular.Font.IsBold);
            Assert.True(bold.Font.IsBold);
            Assert.True(italic.Font.IsItalic);
            OfficeFontFace regularFace = Resolve(regular);
            OfficeFontFace boldFace = Resolve(bold);
            OfficeFontFace italicFace = Resolve(italic);
            Assert.Equal(OfficeFontStyle.Regular,
                Assert.IsType<OfficeTrueTypeFont>(regularFace.Program).FaceDescriptor.ToStyle());
            Assert.Equal(OfficeFontStyle.Bold,
                Assert.IsType<OfficeTrueTypeFont>(boldFace.Program).FaceDescriptor.ToStyle());
            Assert.Equal(OfficeFontStyle.Italic,
                Assert.IsType<OfficeTrueTypeFont>(italicFace.Program).FaceDescriptor.ToStyle());
            Assert.NotEqual(regularFace.Program.Fingerprint, boldFace.Program.Fingerprint);
            Assert.NotEqual(regularFace.Program.Fingerprint, italicFace.Program.Fingerprint);
            Assert.NotEqual(boldFace.Program.Fingerprint, italicFace.Program.Fingerprint);
            rendered.RequireNoLoss();

            OfficeFontFace Resolve(HtmlRenderText run) {
                Assert.True(rendered.Fonts.TryResolveFaceForText(run.Text, run.Font.FamilyName,
                    run.FontDescriptor, run.Font.Size, out OfficeFontFace? face));
                Assert.Equal(run.FontDescriptor, face!.Descriptor);
                return face;
            }
        }
    }

    private static bool HasDistinctHostFaces(string family) {
        OfficeTrueTypeFont? regular = OfficeTrueTypeFont.TryLoadFontFamilyForText(family,
            OfficeFontFaceDescriptor.Regular, "REGULAR", out OfficeFontStyle regularStyle);
        OfficeTrueTypeFont? bold = OfficeTrueTypeFont.TryLoadFontFamilyForText(family,
            OfficeFontFaceDescriptor.FromStyle(OfficeFontStyle.Bold), "BOLD", out OfficeFontStyle boldStyle);
        OfficeTrueTypeFont? italic = OfficeTrueTypeFont.TryLoadFontFamilyForText(family,
            OfficeFontFaceDescriptor.FromStyle(OfficeFontStyle.Italic), "ITALIC", out OfficeFontStyle italicStyle);
        return regular != null && bold != null && italic != null
            && regularStyle == OfficeFontStyle.Regular && boldStyle == OfficeFontStyle.Bold
            && italicStyle == OfficeFontStyle.Italic
            && ((IOfficeFontProgram)regular).Fingerprint != ((IOfficeFontProgram)bold).Fingerprint
            && ((IOfficeFontProgram)regular).Fingerprint != ((IOfficeFontProgram)italic).Fingerprint
            && ((IOfficeFontProgram)bold).Fingerprint != ((IOfficeFontProgram)italic).Fingerprint;
    }
}
