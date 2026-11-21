using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    private const string ShearedImagePlacement = "q 40 0 20 40 30 30 cm /Im1 Do Q";
    private const string ShearedType3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 240 200] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500]";
    // At 1000 pt the glyph space of ShearedType3Font is the page space.
    private const string ShearedType3Show = "BT /FType3 1000 Tf (A) Tj ET";

    // A 40 pt square sheared right by half its height paints a 1600 sq pt parallelogram inside a 60 x 40 pt box,
    // half red and half blue. The stencil mask paints half its columns: an 800 sq pt parallelogram 40 pt wide.
    [Theory]
    [InlineData("type3-image", 1600, 60)]
    [InlineData("type3-image-mask", 800, 40)]
    [InlineData("type3-group-image", 1600, 60)]
    [InlineData("type3-pattern-cell-image", 1600, 60)]
    [InlineData("type3-pattern-image-mask", 800, 40)]
    public void RenderPage_PaintsShearedType3ImageAsItsParallelogram(string context, int area, int width) {
        byte[] pdf = BuildShearedImagePdf(context);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);

        OfficeRasterImage page = OfficeRasterImageDecoder.Decode(PdfPageImageRenderer.RenderPageAsPng(pdf, 1, 1D));
        // A pixel counts as painted when it is at least half covered; interpolation blends the red and blue halves.
        int red = 0, blue = 0, left = int.MaxValue, right = -1, top = int.MaxValue, bottom = -1;
        for (int y = 0; y < page.Height; y++) {
            for (int x = 0; x < page.Width; x++) {
                OfficeColor color = page.GetPixel(x, y);
                if (color.R + color.G + color.B > 510) continue;
                if (color.R >= color.B) red++; else blue++;
                left = Math.Min(left, x); right = Math.Max(right, x); top = Math.Min(top, y); bottom = Math.Max(bottom, y);
            }
        }

        Assert.InRange(red + blue, area * 0.93D, area * 1.07D);
        if (context.EndsWith("-mask", StringComparison.Ordinal)) Assert.Equal(0, blue);
        else Assert.InRange(red, (red + blue) * 0.4D, (red + blue) * 0.6D);
        Assert.InRange(right - left + 1, width - 2, width + 2);
        Assert.InRange(bottom - top + 1, 38, 42);
    }

    // A shear far below a pixel is still outside the tolerance that rounds a matrix to an axis-aligned box,
    // so an image the reader checks as exactly drawable keeps its full matrix instead of being projected.
    [Theory]
    [InlineData("type3-image")]
    [InlineData("type3-group-image")]
    [InlineData("type3-pattern-cell-image")]
    public void RenderPageAsSvg_KeepsTheFullMatrixOfANearlyAxisAlignedType3Image(string context) {
        byte[] pdf = BuildShearedImagePdf(context, "q 40 0 0.0009 40 30 30 cm /Im1 Do Q");

        string svg = System.Text.Encoding.UTF8.GetString(PdfPageImageRenderer.RenderPageAsSvg(pdf));

        Assert.Contains("0.0009", svg, StringComparison.Ordinal);
    }

    private static byte[] BuildShearedImagePdf(string context, string placementContent = ShearedImagePlacement) {
        const string Resources = "<< /Font << /FType3 5 0 R >> >>";
        string image = BuildShearedTestImage(imageMask: false), mask = BuildShearedTestImage(imageMask: true);
        string group = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Group << /Type /Group /S /Transparency /I true /CS /DeviceRGB >> /Resources << /XObject << /Im1 7 0 R >> >>", placementContent);
        string imagePattern = BuildStreamObject(9, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 240 200] /XStep 240 /YStep 200 /Resources << /XObject << /Im1 7 0 R >> >>", placementContent);
        string redPattern = BuildStreamObject(9, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << >>", "1 0 0 rg 0 0 10 10 re f");
        return context switch {
            "type3-image" => BuildSingleStreamPdf(ShearedType3Show, Resources,
                ShearedType3Font + " /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj", BuildStreamObject(6, "<<", "500 0 d0 " + placementContent), image),
            "type3-image-mask" => BuildSingleStreamPdf("1 0 0 rg " + ShearedType3Show, Resources,
                ShearedType3Font + " /PaintType 2 /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj", BuildStreamObject(6, "<<", "500 0 d0 " + placementContent), mask),
            "type3-group-image" => BuildSingleStreamPdf(ShearedType3Show, Resources,
                ShearedType3Font + " /Resources << /XObject << /Fm1 8 0 R >> >> >>\nendobj", BuildStreamObject(6, "<<", "500 0 d0 /Fm1 Do"), image, group),
            "type3-pattern-cell-image" => BuildSingleStreamPdf(ShearedType3Show, Resources,
                ShearedType3Font + " /Resources << /Pattern << /P1 9 0 R >> >> >>\nendobj", BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 240 200 re f"), image, imagePattern),
            "type3-pattern-image-mask" => BuildSingleStreamPdf("/Pattern cs /P1 scn " + ShearedType3Show, "<< /Font << /FType3 5 0 R >> /Pattern << /P1 9 0 R >> >>",
                ShearedType3Font + " /PaintType 2 /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj", BuildStreamObject(6, "<<", "500 0 d0 " + placementContent), mask, redPattern),
            _ => throw new ArgumentOutOfRangeException(nameof(context), context, null)
        };
    }

    // Object 7: a 10x10 image with red left columns and blue right columns, or a stencil mask that
    // marks either its left or its right five columns, depending on how the stencil polarity is read.
    private static string BuildShearedTestImage(bool imageMask) {
        if (imageMask) {
            return BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 10 /Height 10 /ImageMask true /BitsPerComponent 1 /Filter /ASCIIHexDecode",
                string.Concat(Enumerable.Repeat("F800", 10)) + ">");
        }
        string row = string.Concat(Enumerable.Repeat("FF0000", 5)) + string.Concat(Enumerable.Repeat("0000FF", 5));
        // Images inside Type 3 transparency groups are projected only when they declare interpolation.
        return BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 10 /Height 10 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Interpolate true /Filter /ASCIIHexDecode",
            string.Concat(Enumerable.Repeat(row, 10)) + ">");
    }
}
