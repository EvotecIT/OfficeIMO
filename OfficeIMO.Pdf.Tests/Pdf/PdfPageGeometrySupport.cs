using System.Text;

namespace OfficeIMO.Tests.Pdf;

internal static class PdfPageGeometrySupport {
    internal static byte[] BuildPageGeometryPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /MediaBox [0 0 400 300] /BleedBox [5 10 395 290] /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /CropBox [10 20 390 280] /TrimBox [20 30 380 270] /ArtBox [25 35 375 265] /UserUnit 2 /Tabs /S /Dur 5 /Trans << /S /Fly /D 1.5 /Dm /H /M /I /Di 90 /SS 0.75 /B true >> /Metadata 5 0 R /PieceInfo << /OfficeIMO << /LastModified (D:20260607000000Z) >> >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Metadata /Subtype /XML /Length 7 >>",
            "stream",
            "xmpdata",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }
}
