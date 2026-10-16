using System;
using System.Collections.Generic;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfRendererFidelityTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void TransientType0PasswordAppearanceEncodesOnlyMaskCharacters(bool includesMaskGlyph) {
        string mapping = "1 beginbfchar\n<0041> <0041>\nendbfchar\n" +
            (includesMaskGlyph ? "1 beginbfchar\n<002A> <002A>\nendbfchar" : "");
        PdfReadDocument document = CreateTransientType0TextWidget("Identity-H", "(AA)", 8192, 0,
            mapping, "[42 [500] 65 [600]]");
        bool created = PdfFormFiller.TryCreateTransientWidgetAppearance(document.Objects,
            (PdfDictionary)document.Objects[5].Value, (PdfDictionary)document.Objects[8].Value, out PdfStream appearance);
        Assert.Equal(includesMaskGlyph, created);
        if (created) {
            string content = Encoding.ASCII.GetString(appearance.Data);
            Assert.Contains("<002A002A>", content);
            Assert.DoesNotContain("00410041", content);
        }
    }

    [Theory]
    [InlineData(1, "36.4")]
    [InlineData(2, "69.8")]
    public void TransientPredefinedCMapUsesCidWidthsForAlignment(int alignment, string expectedX) {
        PdfReadDocument document = CreateTransientType0TextWidget("UniJIS-UCS2-H", "(A)", 0, alignment,
            "1 beginbfchar\n<0041> <0041>\nendbfchar", "[34 [600] 65 [200]]");
        Assert.True(PdfFormFiller.TryCreateTransientWidgetAppearance(document.Objects,
            (PdfDictionary)document.Objects[5].Value, (PdfDictionary)document.Objects[8].Value, out PdfStream appearance));
        Assert.Contains(expectedX + " ", Encoding.ASCII.GetString(appearance.Data));
    }

    private static PdfReadDocument CreateTransientType0TextWidget(string encoding, string value, int flags, int alignment,
        string mapping, string widths) {
        byte[] cmap = Encoding.ASCII.GetBytes(mapping);
        return PdfReadDocument.Open(BuildPdf("", "<< >>", objects: new List<(string, byte[])> {
            ($"<< /Type /Annot /Subtype /Widget /FT /Tx /Rect [10 60 90 80] /DA (/F1 12 Tf) /Ff {flags} /Q {alignment} /V {value} >>", Array.Empty<byte>()),
            ($"<< /Type /Font /Subtype /Type0 /BaseFont /Custom /Encoding /{encoding} /ToUnicode 7 0 R /DescendantFonts [<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Custom /CIDSystemInfo << /Registry (Adobe) /Ordering (Japan1) /Supplement 6 >> /DW 1000 /W {widths} >>] >>", Array.Empty<byte>()),
            ($"<< /Length {cmap.Length} >>", cmap),
            ("<< /DR << /Font << /F1 6 0 R >> >> >>", Array.Empty<byte>())
        }));
    }

    [Fact]
    public void MeshBackgroundCoversTheClipOutsideItsTriangles() {
        var objects = new List<(string, byte[])> {
            ("<< /ShadingType 4 /ColorSpace /DeviceRGB /BitsPerCoordinate 8 /BitsPerComponent 8 " +
             "/BitsPerFlag 8 /Decode [0 100 0 100 0 1 0 1 0 1] /Background [0 1 0] /Length 18 >>",
                new byte[] { 0, 25, 25, 255, 0, 0, 0, 128, 25, 255, 0, 0, 0, 25, 128, 255, 0, 0 })
        };
        OfficeRasterImage image = Render(BuildPdf("/Sh1 sh", "<< /Shading << /Sh1 5 0 R >> >>", objects: objects));
        AssertColor(image, 80, 20, 0, 255, 0);
        AssertColor(image, 20, 80, 255, 0, 0);
    }

    [Fact]
    public void MissingPushButtonAppearanceUsesItsCaption() {
        PdfReadDocument document = PdfReadDocument.Open(BuildPdf("", "<< >>", objects: new List<(string, byte[])> {
            ("<< /Type /Annot /Subtype /Widget /FT /Btn /Ff 65536 /Rect [10 60 90 80] /MK << /CA (Submit) >> >>", Array.Empty<byte>())
        }));
        Assert.True(PdfFormFiller.TryCreateTransientWidgetAppearance(document.Objects,
            (PdfDictionary)document.Objects[5].Value, null, out PdfStream appearance));
        Assert.Contains("5375626D6974", Encoding.ASCII.GetString(appearance.Data));
    }

    [Fact]
    public void AppearanceLessRadioGroupSelectsOnlyTheMatchingExportOption() {
        var objects = new List<(string, byte[])> {
            ("<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [10 60 30 80] /P 3 0 R >>", Array.Empty<byte>()),
            ("<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [50 60 70 80] /P 3 0 R >>", Array.Empty<byte>()),
            ("<< /FT /Btn /Ff 32768 /V (Two) /Opt [(One) (Two)] /Kids [5 0 R 6 0 R] >>", Array.Empty<byte>())
        };
        OfficeRasterImage image = Render(BuildPdf("", "<< >>", objects: objects,
            catalogEntries: "/AcroForm << /Fields [7 0 R] >>", pageEntries: "/Annots [5 0 R 6 0 R]"));
        AssertColor(image, 20, 30, 255, 255, 255);
        AssertColor(image, 60, 30, 0, 0, 0);
    }

    [Theory]
    [InlineData("/Subtype /Type1 /BaseFont /Helvetica /Encoding << /BaseEncoding /WinAnsiEncoding /Differences [65 /B] >>", "(A)", false)]
    [InlineData("/Subtype /Type0 /BaseFont /Custom /Encoding /Identity-H /ToUnicode 7 0 R", "<FEFF03A9>", true)]
    public void TransientTextUsesAnEncoderThatMatchesItsFont(string font, string value, bool unicode) {
        byte[] cmap = Encoding.ASCII.GetBytes("1 beginbfchar\n<1234> <03A9>\nendbfchar");
        var objects = new List<(string, byte[])> {
            ("<< /Type /Annot /Subtype /Widget /FT /Tx /Rect [10 60 90 80] /DA (/F1 12 Tf) /V " + value + " >>", Array.Empty<byte>()),
            ("<< /Type /Font " + font + " >>", Array.Empty<byte>()),
            ("<< /Length " + cmap.Length + " >>", cmap)
        };
        PdfReadDocument document = PdfReadDocument.Open(BuildPdf("", "<< >>", objects: objects));
        var acroForm = new PdfDictionary();
        var fonts = new PdfDictionary(); fonts.Items["F1"] = new PdfReference(6, 0);
        var resources = new PdfDictionary(); resources.Items["Font"] = fonts; acroForm.Items["DR"] = resources;
        Assert.True(PdfFormFiller.TryCreateTransientWidgetAppearance(document.Objects,
            (PdfDictionary)document.Objects[5].Value, acroForm, out PdfStream appearance));
        string content = Encoding.ASCII.GetString(appearance.Data);
        Assert.Contains(unicode ? "/F1 " : "/Helv ", content);
        Assert.Contains(unicode ? "<1234>" : "<41>", content);
        if (unicode) {
            Assert.Throws<PdfReadLimitException>(() => PdfFormFiller.TryCreateTransientWidgetAppearance(document.Objects,
                (PdfDictionary)document.Objects[5].Value, acroForm, out _, maximumDecodedBytes: 8));
        }
    }
}
