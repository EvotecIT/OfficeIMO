using OfficeIMO.Pdf;
using System.Globalization;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfMarkupAnnotationTextTests {
    // "Alpha Bravo Charlie" in 12 pt Helvetica at x = 72: by the Helvetica metrics "Bravo" spans x 106.02 to 137.36
    // and "Charlie" x 140.70 to 178.70. The quads below are written from those metrics, not from OfficeIMO's reading.
    private const string BravoHighlight =
        "<< /Type /Annot /Subtype /Highlight /Rect [100 690 140 715] /T (Alice) /Contents (Check this) /QuadPoints [105 712 138 712 105 696 138 696] >>";

    [Fact]
    public void GetMarkupText_SelectsGlyphsWhoseCentersLieInsideEachQuad() {
        byte[] pdf = BuildPdf(unmappableRun: false,
            // Viewer corner order: top-left, top-right, bottom-left, bottom-right. Only "Bravo" lies inside.
            BravoHighlight,
            // Specification corner order (counterclockwise from bottom-left), spanning two words.
            "<< /Type /Annot /Subtype /Squiggly /Rect [100 690 190 715] /QuadPoints [105 696 190 696 190 712 105 712] >>",
            // A degenerate quad covers nothing, so /Rect (around "Charlie") applies.
            "<< /Type /Annot /Subtype /Underline /Rect [139 696 180 712] /QuadPoints [150 700 150 700 150 700 150 700] >>",
            "<< /Type /Annot /Subtype /Text /Rect [72 650 90 668] /Contents (Sticky note) >>");

        PdfLogicalPage page = Assert.Single(PdfDocument.Load(new MemoryStream(pdf, writable: false)).Read().Pages);
        PdfLogicalPage otherRead = Assert.Single(PdfDocument.Load(new MemoryStream(pdf, writable: false)).Read().Pages);

        Assert.Equal(4, page.Annotations.Count);
        Assert.Equal("Bravo", page.GetMarkupText(page.Annotations[0]));
        Assert.Equal("Bravo Charlie", page.GetMarkupText(page.Annotations[1]));
        Assert.Equal("Charlie", page.GetMarkupText(page.Annotations[2]));
        Assert.Null(page.GetMarkupText(page.Annotations[3]));
        // An annotation read from another document is not this page's annotation, even on the same page number.
        Assert.Null(page.GetMarkupText(otherRead.Annotations[0]));
    }

    [Fact]
    public void GetMarkupText_LeavesOutUnmappableRunsOnARecoveredPage() {
        byte[] pdf = BuildPdf(unmappableRun: true, BravoHighlight);

        PdfInspectionReadResult inspection = PdfDocument.Load(new MemoryStream(pdf, writable: false))
            .ReadForInspection(new PdfLoadOptions { RecoverDecodableTextForInspection = true });
        PdfLogicalPage page = Assert.Single(inspection.Document.Pages);

        Assert.Equal("Bravo", page.GetMarkupText(page.Annotations[0]));
    }

    private static byte[] BuildPdf(bool unmappableRun, params string[] annotations) {
        string content = "BT /F1 12 Tf 72 700 Td (Alpha Bravo Charlie) Tj ET" +
            (unmappableRun ? " BT /F2 12 Tf 72 600 Td <00010002> Tj ET" : string.Empty);
        int firstAnnotation = 7;
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 800] /Resources << /Font << /F1 5 0 R /F2 6 0 R >> >> /Contents 4 0 R /Annots ["
                + string.Join(" ", annotations.Select((_, index) => (index + firstAnnotation).ToString(CultureInfo.InvariantCulture) + " 0 R")) + "] >>",
            "<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + content + "\nendstream",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            // A predefined CMap name OfficeIMO cannot map to Unicode.
            "<< /Type /Font /Subtype /Type0 /BaseFont /Unmapped /Encoding /Mystery-H >>"
        };
        objects.AddRange(annotations);

        var builder = new StringBuilder("%PDF-1.4\n");
        var offsets = new List<int>();
        for (int i = 0; i < objects.Count; i++) {
            offsets.Add(builder.Length);
            builder.Append(i + 1).Append(" 0 obj\n").Append(objects[i]).Append("\nendobj\n");
        }
        int xref = builder.Length;
        builder.Append("xref\n0 ").Append(objects.Count + 1).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets) builder.Append(offset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        builder.Append("trailer\n<< /Root 1 0 R /Size ").Append(objects.Count + 1).Append(" >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }
}
