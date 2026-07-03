using System.Text;

namespace OfficeIMO.Tests.Pdf;

internal static class PdfOptionalContentSupport {
    internal static byte[] BuildOptionalContentMetadataPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.5",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [5 0 R 6 0 R] /D << /Name (Default layers) /Creator (OfficeIMO fixture) /BaseState /ON /ON [5 0 R] /OFF [6 0 R] /Locked [6 0 R] /Order [(Layers) [5 0 R 6 0 R]] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /OCG /Name (Print layer) /Intent [/View /Design] /Usage << /CreatorInfo << /Creator (OfficeIMO) /Subtype /Artwork >> /View << /ViewState /ON >> /Print << /PrintState /ON >> /Export << /ExportState /OFF >> >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /OCG /Name (Hidden layer) /Intent /View /Usage << /View << /ViewState /OFF >> /Print << /PrintState /OFF >> /Export << /ExportState /ON >> >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }
}
