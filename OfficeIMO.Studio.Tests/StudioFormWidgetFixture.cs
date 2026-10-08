namespace OfficeIMO.Studio.Tests;

/// <summary>One shared text field with two widgets on page one and one on a rotated page.</summary>
internal static class StudioFormWidgetFixture {
    internal static string SharedTextWidgets() => string.Join("\n", new[] {
        "%PDF-1.7",
        "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [5 0 R] >> >>", "endobj",
        "2 0 obj", "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>", "endobj",
        "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Annots [6 0 R 7 0 R] >>", "endobj",
        "4 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Rotate 90 /Annots [8 0 R] >>", "endobj",
        "5 0 obj", "<< /FT /Tx /T (Shared) /V (Value) /Kids [6 0 R 7 0 R 8 0 R] >>", "endobj",
        "6 0 obj", "<< /Type /Annot /Subtype /Widget /Parent 5 0 R /Rect [20 20 100 40] /P 3 0 R /F 4 >>", "endobj",
        "7 0 obj", "<< /Type /Annot /Subtype /Widget /Parent 5 0 R /Rect [120 20 220 40] /P 3 0 R /F 4 >>", "endobj",
        "8 0 obj", "<< /Type /Annot /Subtype /Widget /Parent 5 0 R /Rect [20 20 100 40] /P 4 0 R /F 4 >>", "endobj",
        "trailer", "<< /Root 1 0 R /Size 9 >>", "%%EOF"
    }) + "\n";
}
