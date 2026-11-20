using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    // Two L-shapes whose intersection is the top-left and bottom-right corner squares of their shared bounds.
    private const string LowerLeftLClip = "20 20 m 220 20 l 220 60 l 60 60 l 60 180 l 20 180 l h W n\n";
    private const string UpperRightLClip = "220 180 m 20 180 l 20 140 l 180 140 l 180 20 l 220 20 l h W n\n";

    [Fact]
    public void RenderPage_PaintsOnlyTheTrueIntersectionOfTwoNonConvexClips() {
        byte[] pdf = BuildSingleStreamPdf(LowerLeftLClip + UpperRightLClip + "1 0 0 rg 0 0 240 200 re f");

        OfficeRasterImage page = OfficeRasterImageDecoder.Decode(PdfPageImageRenderer.RenderPageAsPng(pdf, 1, 1D));

        Assert.Equal(OfficeColor.Red, page.GetPixel(40, 40));
        Assert.Equal(OfficeColor.Red, page.GetPixel(200, 160));
        Assert.Equal(OfficeColor.White, page.GetPixel(120, 100));
        Assert.Equal(OfficeColor.White, page.GetPixel(40, 100));
        Assert.Equal(OfficeColor.White, page.GetPixel(120, 40));
        Assert.InRange(CountNonWhitePixels(page), 3000, 3400);
    }

    [Fact]
    public void RenderPage_HonoursEachFillRuleWhenIntersectingMultiContourClips() {
        // Even-odd ring with a hole, then a nonzero band whose inner contour must not punch a hole.
        byte[] pdf = BuildSingleStreamPdf(
            "20 20 m 220 20 l 220 180 l 20 180 l h 80 60 m 160 60 l 160 140 l 80 140 l h W* n\n" +
            "0 90 m 240 90 l 240 110 l 0 110 l h 30 95 m 70 95 l 70 105 l 30 105 l h W n\n" +
            "1 0 0 rg 0 0 240 200 re f");

        OfficeRasterImage page = OfficeRasterImageDecoder.Decode(PdfPageImageRenderer.RenderPageAsPng(pdf, 1, 1D));

        Assert.Equal(OfficeColor.Red, page.GetPixel(50, 100));
        Assert.Equal(OfficeColor.Red, page.GetPixel(190, 100));
        Assert.Equal(OfficeColor.White, page.GetPixel(120, 100));
        Assert.Equal(OfficeColor.White, page.GetPixel(50, 80));
    }

    [Fact]
    public void RenderPage_ClipsImageToTheTrueIntersectionOfTwoNonConvexClips() {
        string image = BuildStreamObject(5, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8", "x");
        byte[] pdf = BuildSingleStreamPdf(
            LowerLeftLClip + UpperRightLClip + "q 240 0 0 200 0 0 cm /Im1 Do Q",
            "<< /XObject << /Im1 5 0 R >> >>",
            image);

        OfficeRasterImage page = OfficeRasterImageDecoder.Decode(PdfPageImageRenderer.RenderPageAsPng(pdf, 1, 1D));

        Assert.NotEqual(OfficeColor.White, page.GetPixel(40, 40));
        Assert.NotEqual(OfficeColor.White, page.GetPixel(200, 160));
        Assert.Equal(OfficeColor.White, page.GetPixel(120, 100));
        Assert.Equal(OfficeColor.White, page.GetPixel(40, 100));
        Assert.InRange(CountNonWhitePixels(page), 3000, 3400);
    }

    [Fact]
    public void RenderPage_ClipsTextToTheTrueIntersectionOfTwoNonConvexClips() {
        byte[] pdf = BuildSingleStreamPdf(
            LowerLeftLClip + UpperRightLClip +
            "0 0 0 rg BT /F1 30 Tf 90 90 Td (HH) Tj ET\n" +
            "BT /F1 30 Tf 24 148 Td (H) Tj ET",
            EmptyClipFontResources,
            EmptyClipFontObject);

        OfficeRasterImage page = OfficeRasterImageDecoder.Decode(PdfPageImageRenderer.RenderPageAsPng(pdf, 1, 1D));

        Assert.Equal(0, CountNonWhitePixels(page, 60, 60, 180, 140));
        Assert.True(CountNonWhitePixels(page, 20, 20, 60, 60) > 0);
    }

    [Fact]
    public void RenderPageAsSvg_ExpressesTheIntersectionOfTwoNonConvexClips() {
        byte[] pdf = BuildSingleStreamPdf(LowerLeftLClip + UpperRightLClip + "1 0 0 rg 0 0 240 200 re f");

        string svg = Encoding.UTF8.GetString(PdfPageImageRenderer.RenderPageAsSvg(pdf));

        // The group clip is chained through clipPath elements that each hold one L-shape in the group's local space.
        string groupClipId = Regex.Match(svg, "<g clip-path=\"url\\(#([^\"]+)\\)\"").Groups[1].Value;
        string[] chain = FollowClipPathChain(svg, groupClipId);
        Assert.Equal(3, chain.Length);
        Assert.StartsWith("<rect x=\"0\" y=\"0\" width=\"200\" height=\"160\"", chain[0], StringComparison.Ordinal);
        Assert.Equal(
            new[] { "M0 160L200 160L200 120L40 120L40 0L0 0Z", "M200 0L0 0L0 40L160 40L160 160L200 160Z" },
            chain.Skip(1).Select(content => Regex.Match(content, "<path d=\"([^\"]+)\"").Groups[1].Value).OrderBy(d => d, StringComparer.Ordinal).ToArray());
    }

    private static string[] FollowClipPathChain(string svg, string id) {
        var contents = new List<string>();
        while (id.Length > 0 && contents.Count < 16) {
            Match clipPath = Regex.Match(svg, "<clipPath id=\"" + Regex.Escape(id) + "\"(?: clip-path=\"url\\(#([^\"]+)\\)\")?>(.*?)</clipPath>");
            Assert.True(clipPath.Success, svg);
            contents.Add(clipPath.Groups[2].Value);
            id = clipPath.Groups[1].Value;
        }
        return contents.ToArray();
    }

    private static int CountNonWhitePixels(OfficeRasterImage image) =>
        CountNonWhitePixels(image, 0, 0, image.Width, image.Height);

    private static int CountNonWhitePixels(OfficeRasterImage image, int left, int top, int right, int bottom) {
        int count = 0;
        for (int y = top; y < bottom; y++) {
            for (int x = left; x < right; x++) {
                if (image.GetPixel(x, y) != OfficeColor.White) count++;
            }
        }
        return count;
    }
}
