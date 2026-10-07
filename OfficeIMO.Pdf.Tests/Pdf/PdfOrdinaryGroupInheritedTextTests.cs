using System;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalDocumentCompatibilityTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IsolatedGroupPreservesInheritedTextStateAndShadowedFont(bool nested) {
        byte[] source = InheritedTextGroupFixture(true, nested);
        var page = PdfReadDocument.Open(source).Pages[0];
        Assert.Contains("Alpha", page.ExtractText());
        OfficeDrawing opaque = PdfReadDocument.Open(InheritedTextGroupFixture(false, nested)).Pages[0].ToDrawing();
        var expected = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(140, 80)
            .AddEffectDrawing(opaque, OfficeTransform.Identity, .5), background: OfficeColor.White);
        var actual = OfficeDrawingRasterRenderer.Render(page.ToDrawing(), background: OfficeColor.White);
        int painted = 0;
        for (int y = 0; y < 80; y++) for (int x = 0; x < 140; x++) {
            OfficeColor a = actual.GetPixel(x, y), e = expected.GetPixel(x, y);
            Assert.InRange(Math.Abs(a.R-e.R), 0, 3); Assert.InRange(Math.Abs(a.G-e.G), 0, 3);
            Assert.InRange(Math.Abs(a.B-e.B), 0, 3);
            if (a.R < 240) painted++;
        }
        Assert.True(painted > 100);
    }

    private static byte[] InheritedTextGroupFixture(bool isolated, bool nested) {
        string Stream(int id, string content, string extra = "") =>
            BuildStreamObject(id, Encoding.ASCII.GetBytes(content), extra);
        string group = isolated ? " /Group << /S /Transparency /I true >>" : "";
        string invoke = nested ? "/Outer Do" : "/Fm Do";
        return BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 140 80] /Resources << /Font << /F1 4 0 R >> /ExtGState << /A << /ca " + (isolated ? "0.5" : "1") + " >> >> /XObject << /Fm 6 0 R /Outer 8 0 R >> >> /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            Stream(5, "/F1 20 Tf 1 Tc 80 Tz 3 Ts /A gs " + invoke),
            Stream(6, "BT 15 20 Td (Alpha) Tj ET", "/Type /XObject /Subtype /Form /BBox [0 0 140 80] /Resources << /Font << /F1 7 0 R >> >>" + group),
            "7 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Symbol >>\nendobj",
            Stream(8, "/Fm Do", "/Type /XObject /Subtype /Form /BBox [0 0 140 80] /Resources << /XObject << /Fm 6 0 R >> >>")
        }, rootObjectNumber: 1);
    }
}
