using System.Text;

namespace OfficeIMO.Tests.Pdf;

internal static class PdfAssociatedFileTestSupport {
    internal const string Payload = "PAGE-ASSOCIATED-FILE-PAYLOAD";

    internal static byte[] BuildPageAssociatedFilePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /AF [5 0 R] >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "5 0 obj", "<< /Type /Filespec /F (page.txt) /UF (page.txt) /AFRelationship /Data /EF << /F 6 0 R /UF 6 0 R >> >>", "endobj",
            "6 0 obj", "<< /Type /EmbeddedFile /Subtype /text#2Fplain /Length " + Payload.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>", "stream", Payload, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    internal static byte[] BuildFileAttachmentAnnotationPdf() {
        return BuildFileAttachmentAnnotationPdf(includePageAssociatedFile: true);
    }

    internal static byte[] BuildStandaloneFileAttachmentAnnotationPdf() {
        return BuildFileAttachmentAnnotationPdf(includePageAssociatedFile: false);
    }

    internal static byte[] BuildDuplicateNamedFileAttachmentAnnotationsPdf() {
        const string firstPayload = "FIRST-DUPLICATE-PAYLOAD";
        const string secondPayload = "SECOND-DUPLICATE-PAYLOAD";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [7 0 R 10 0 R] >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "5 0 obj", "<< /Type /Filespec /F (duplicate.txt) /UF (duplicate.txt) /EF << /F 6 0 R /UF 6 0 R >> >>", "endobj",
            "6 0 obj", "<< /Type /EmbeddedFile /Subtype /text#2Fplain /Length " + firstPayload.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>", "stream", firstPayload, "endstream", "endobj",
            "7 0 obj", "<< /Type /Annot /Subtype /FileAttachment /Rect [10 10 30 30] /FS 5 0 R >>", "endobj",
            "8 0 obj", "<< /Type /Filespec /F (duplicate.txt) /UF (duplicate.txt) /EF << /F 9 0 R /UF 9 0 R >> >>", "endobj",
            "9 0 obj", "<< /Type /EmbeddedFile /Subtype /text#2Fplain /Length " + secondPayload.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>", "stream", secondPayload, "endstream", "endobj",
            "10 0 obj", "<< /Type /Annot /Subtype /FileAttachment /Rect [40 10 60 30] /FS 8 0 R >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 11 >>", "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFileAttachmentAnnotationPdf(bool includePageAssociatedFile) {
        string associatedFiles = includePageAssociatedFile ? " /AF [5 0 R]" : string.Empty;
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R" + associatedFiles + " /Annots [7 0 R] >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "5 0 obj", "<< /Type /Filespec /F (page.txt) /UF (page.txt) /AFRelationship /Data /EF << /F 6 0 R /UF 6 0 R >> >>", "endobj",
            "6 0 obj", "<< /Type /EmbeddedFile /Subtype /text#2Fplain /Length " + Payload.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>", "stream", Payload, "endstream", "endobj",
            "7 0 obj", "<< /Type /Annot /Subtype /FileAttachment /Rect [10 10 30 30] /FS 5 0 R >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 8 >>", "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }
}
