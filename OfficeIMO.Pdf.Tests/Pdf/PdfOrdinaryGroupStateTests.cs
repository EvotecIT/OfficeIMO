using System;
using System.IO;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalDocumentCompatibilityTests {
    [Theory]
    [InlineData("text", true)]
    [InlineData("image", true)]
    [InlineData("path", true)]
    [InlineData("text", false)]
    [InlineData("image", false)]
    [InlineData("path", false)]
    public void HiddenOptionalContentFormsRemainUnpainted(string paint, bool isolated) {
        string content = paint switch {
            "text" => "BT /F1 12 Tf 5 5 Td (Hidden) Tj ET",
            "image" => "10 0 0 10 5 5 cm BI /W 1 /H 1 /CS /RGB /BPC 8 ID ABC EI",
            _ => "1 0 0 rg 5 5 10 10 re f"
        };
        byte[] pdf = BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [7 0 R] /D << /BaseState /ON /OFF [7 0 R] >> >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 40 40] /Resources << /Font << /F1 4 0 R >> /XObject << /Fm 6 0 R >> >> /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("/Fm Do")),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(content), "/Type /XObject /Subtype /Form /BBox [0 0 40 40] /OC 7 0 R" + (isolated ? " /Group << /S /Transparency /I true >>" : "")),
            "7 0 obj\n<< /Type /OCG /Name (Hidden form) >>\nendobj"
        }, rootObjectNumber: 1);
        var page = PdfReadDocument.Open(pdf).Pages[0];
        Assert.DoesNotContain("Hidden", page.ExtractText());
        var image = OfficeDrawingRasterRenderer.Render(page.ToDrawing(), background: OfficeColor.White);
        for(int y=0;y<40;y++) for(int x=0;x<40;x++) Assert.Equal(OfficeColor.White,image.GetPixel(x,y));
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(true, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void GroupInheritsFontSelectedThroughExtGState(bool declared, bool nested, bool restore) {
        byte[] Fixture(bool isolated) => BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 140 80] /Resources << " + (declared ? "/Font << /F2 4 0 R /F3 7 0 R >> " : "") + "/ExtGState << /A << /Font [4 0 R 30] /ca " + (isolated ? ".5" : "1") + " >> /B << /Font [7 0 R 12] >> >> /XObject << /Fm 6 0 R /Outer 8 0 R >> >> /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5,Encoding.ASCII.GetBytes("/A gs " + (restore ? "q /B gs Q " : "") + (nested ? "/Outer Do" : "/Fm Do"))),
            BuildStreamObject(6,Encoding.ASCII.GetBytes("BT 15 20 Td (Alpha) Tj ET"),"/Type /XObject /Subtype /Form /BBox [0 0 140 80] /Resources << >>"+(isolated?" /Group << /S /Transparency /I true >>":"")),
            "7 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Symbol >>\nendobj",
            BuildStreamObject(8,Encoding.ASCII.GetBytes("/Fm Do"),"/Type /XObject /Subtype /Form /BBox [0 0 140 80] /Resources << /XObject << /Fm 6 0 R >> >>")
        },rootObjectNumber:1);
        var opaque=PdfReadDocument.Open(Fixture(false)).Pages[0].ToDrawing();
        var expected=OfficeDrawingRasterRenderer.Render(new OfficeDrawing(140,80).AddEffectDrawing(opaque,OfficeTransform.Identity,.5),background:OfficeColor.White);
        var actual=OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(Fixture(true)).Pages[0].ToDrawing(),background:OfficeColor.White);
        for(int y=0;y<80;y++)for(int x=0;x<140;x++) {
            var e=expected.GetPixel(x,y);var a=actual.GetPixel(x,y);
            Assert.InRange(Math.Abs(e.R-a.R),0,3);Assert.InRange(Math.Abs(e.G-a.G),0,3);Assert.InRange(Math.Abs(e.B-a.B),0,3);
        }
    }
}
