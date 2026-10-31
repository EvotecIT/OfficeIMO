using System.Collections.Generic;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalDocumentCompatibilityTests {
    [Theory]
    [InlineData(2)]
    [InlineData(8)]
    public void AcyclicSoftMaskChainsRespectTheConfiguredNestingLimit(int limit) {
        const int masks = 3;
        var objects = new List<string> {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 20 20] /Resources << /XObject << /G 5 0 R >> >> /Contents 4 0 R >>\nendobj",
            BuildStreamObject(4, Encoding.ASCII.GetBytes("/G Do"))
        };
        for (int index = 0; index <= masks; index++) {
            bool masked = index < masks;
            string resources = masked ? "/Resources << /ExtGState << /A << /SMask << /S /Alpha /G " + (index + 6) + " 0 R >> >> >> >>" : "";
            objects.Add(BuildStreamObject(index + 5, Encoding.ASCII.GetBytes((masked ? "/A gs " : "") + "1 g 0 0 20 20 re f"),
                "/Type /XObject /Subtype /Form /BBox [0 0 20 20] /Group << /S /Transparency /I true /CS /DeviceRGB >> " + resources));
        }
        var page = PdfReadDocument.Open(BuildPdf(objects, rootObjectNumber: 1), new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxContentNestingDepth = limit }
        }).Pages[0];

        if (limit == 2) {
            var error = Assert.Throws<PdfReadLimitException>(() => page.ToDrawing());
            Assert.Equal(PdfReadLimitKind.ContentNestingDepth, error.Kind);
            Assert.Equal(2, error.Limit);
            Assert.Equal(3, error.Actual);
        } else {
            Assert.Equal(OfficeColor.White, OfficeDrawingRasterRenderer.Render(page.ToDrawing()).GetPixel(10, 10));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IsolatedGroupsRetainTheActiveSoftMaskAncestry(bool cyclic) {
        byte[] pdf = BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 20 20] /Resources << /XObject << /G 5 0 R >> >> /Contents 4 0 R >>\nendobj",
            BuildStreamObject(4, Encoding.ASCII.GetBytes("/G Do /G Do")),
            BuildStreamObject(5, Encoding.ASCII.GetBytes("/A gs 1 0 0 rg 0 0 20 20 re f"),
                "/Type /XObject /Subtype /Form /BBox [0 0 20 20] /Group << /S /Transparency /I true /CS /DeviceRGB >> " +
                "/Resources << /ExtGState << /A << /SMask << /S /Alpha /G 6 0 R >> >> >> >>"),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(cyclic ? "/G Do" : "1 g 0 0 20 20 re f"),
                "/Type /XObject /Subtype /Form /BBox [0 0 20 20] /Group << /S /Transparency /I true /CS /DeviceRGB >> " +
                "/Resources << /XObject << /G 5 0 R >> >>")
        }, rootObjectNumber: 1);
        // Repeated cached streams must still retain their active resource ancestry.
        var page = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPageContentBytes = 2048 }
        }).Pages[0];

        var image = OfficeDrawingRasterRenderer.Render(page.ToDrawing(), background: OfficeColor.White);

        Assert.Equal(cyclic ? OfficeColor.White : OfficeColor.Red, image.GetPixel(10, 10));
    }
}
