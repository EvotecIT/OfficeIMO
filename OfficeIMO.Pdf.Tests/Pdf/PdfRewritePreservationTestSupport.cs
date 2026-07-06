using System.Text;
using OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

internal static class PdfRewritePreservationTestSupport {
    public static byte[] BuildPreservationProofPdf() {
        var options = new PdfOptions {
                IncludeXmpMetadata = true
            }
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B)
            .SetCatalogView(PdfCatalogPageMode.UseThumbs, PdfCatalogPageLayout.SinglePage)
            .SetLanguage("en-US")
            .AddEmbeddedFile("preservation.txt", Encoding.UTF8.GetBytes("Preservation attachment"), "text/plain", PdfAssociatedFileRelationship.Data);

        return PdfDocument.Create(options)
            .Meta(
                title: "Original preservation title",
                author: "OfficeIMO",
                subject: "Rewrite preservation proof",
                keywords: "pdf,preservation")
            .Paragraph(p => p.Text("PreservationMarker"))
            .Paragraph(p => p.Link("ExternalLink", "https://evotec.xyz"))
            .Paragraph(p => p.LinkToBookmark("JumpToSecondPage", "SecondPage"))
            .PageBreak()
            .Bookmark("SecondPage")
            .Paragraph(p => p.Text("SecondPageMarker"))
            .ToBytes();
    }

    public static byte[] BuildTaggedPreservationProofPdf() {
        return PdfDocument.Create()
            .TaggedPdfCatalogMarkers()
            .Language("en-US")
            .H1("Tagged preservation heading")
            .Paragraph(p => p.Text("TaggedPreservationMarker"))
            .ToBytes();
    }

    public static byte[] BuildNavigationPreservationProofPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageLabels 7 0 R /Names << /Dests << /Names [(Chapter1) [3 0 R /Fit] (Chapter2) [4 0 R /XYZ 10 180 1]] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Nums [0 << /S /r /P (front-) /St 1 >> 1 << /S /D /St 2 >>] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    public static byte[] BuildViewerActionPreservationProofPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OpenAction [3 0 R /XYZ 10 180 0] /ViewerPreferences 6 0 R /Names << /JavaScript << /Names [(Open) 7 0 R] >> >> /AA << /WC 8 0 R >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /AA << /O 9 0 R /C 10 0 R >> >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "6 0 obj",
            "<< /HideToolbar true /DisplayDocTitle true >>",
            "endobj",
            "7 0 obj",
            "<< /S /JavaScript /JS (app.alert('Open')) >>",
            "endobj",
            "8 0 obj",
            "<< /S /Launch /F (tool.exe) >>",
            "endobj",
            "9 0 obj",
            "<< /S /JavaScript /JS (app.alert('Page open')) >>",
            "endobj",
            "10 0 obj",
            "<< /S /Launch /F (page.exe) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 11 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    public static byte[] BuildSourceStructurePreservationProofPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Version /1.7 /Pages 2 0 R >>",
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
            "<< /Type /XRef /W [1 2 1] /Size 6 /Root 1 0 R >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "6 0 obj",
            "<< /Type /ObjStm /N 0 /First 0 /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 /Prev 100 >>",
            "startxref",
            "100",
            "%%EOF",
            "startxref",
            "200",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    public static byte[] BuildSignedIncrementalProofPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 7 0 R /Perms << /DocMDP 6 0 R /UR3 6 0 R >> /DSS 9 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /FT /Sig /T (Approval) /V 6 0 R /Subtype /Widget /Rect [10 10 120 40] /Lock << /Type /SigFieldLock /Action /Include /Fields [(Total) (Approver)] >> /SV << /Filter /Adobe.PPKLite /SubFilter [/adbe.pkcs7.detached] /DigestMethod [/SHA256 /SHA512] /Reasons [(Approval) (Final)] /Ff 3 /AddRevInfo true /MDP << /P 2 >> >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /Name (Alice) /Location (Warsaw) /Reason (Approval) /ContactInfo (alice@example.test) /M (D:20260607120000+02'00') /ByteRange [0 10 20 30] /Contents <001122> /Reference [<< /TransformMethod /DocMDP /TransformParams << /Type /TransformParams /V /1.2 /P 2 >> >>] >>",
            "endobj",
            "7 0 obj",
            "<< /Fields [5 0 R] /SigFlags 3 >>",
            "endobj",
            "8 0 obj",
            "<< /Producer (OfficeIMO signed fixture) >>",
            "endobj",
            "9 0 obj",
            "<< /Certs [10 0 R 11 0 R] /OCSPs [12 0 R] /CRLs [13 0 R] /VRI << /ABCDEF << /Cert [10 0 R] /OCSP [12 0 R] /CRL [13 0 R] /TS 14 0 R >> >> >>",
            "endobj",
            "10 0 obj",
            "<< /Type /EmbeddedFile /Length 0 >>",
            "endobj",
            "11 0 obj",
            "<< /Type /EmbeddedFile /Length 0 >>",
            "endobj",
            "12 0 obj",
            "<< /Type /EmbeddedFile /Length 0 >>",
            "endobj",
            "13 0 obj",
            "<< /Type /EmbeddedFile /Length 0 >>",
            "endobj",
            "14 0 obj",
            "<< /Type /TimestampEvidence /Length 0 >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Info 8 0 R /ID [(abc) (def)] /Size 15 /Prev 100 >>",
            "startxref",
            "100",
            "%%EOF",
            "startxref",
            "200",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }
}
