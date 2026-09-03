using OfficeIMO.ContentSafety;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfHiddenContentInspectionTests {
    [Fact]
    public void ContentSafetySurfacesHiddenLayerTextIncludingHiddenFormContent() {
        byte[] pdf = BuildHiddenOptionalContentPdf();

        string visibleText = PdfDocument.Load(pdf).Read().Text;
        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(pdf);
        OfficeContentSafetyFinding[] hidden = report.Findings
            .Where(finding => finding.Kind == OfficeContentConcealmentKind.HiddenContainer)
            .ToArray();

        Assert.Contains("VISIBLE-CONTENT", visibleText, StringComparison.Ordinal);
        Assert.DoesNotContain("HIDDEN-PAGE-CONTENT", visibleText, StringComparison.Ordinal);
        Assert.DoesNotContain("HIDDEN-FORM-CONTENT", visibleText, StringComparison.Ordinal);
        Assert.Contains(hidden, finding => finding.TextPreview.Contains("HIDDEN-PAGE-CONTENT", StringComparison.Ordinal));
        Assert.Contains(hidden, finding => finding.TextPreview.Contains("HIDDEN-FORM-CONTENT", StringComparison.Ordinal));
        Assert.All(hidden, finding => Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability));
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Contains("default layer configuration", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetySurfacesHiddenAnnotationAndWidgetValues() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 8 0 R >> >> /Contents 4 0 R /Annots [5 0 R 7 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "5 0 obj\n<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (HIDDEN-ANNOTATION) /F 2 >>\nendobj",
            "6 0 obj\n<< /FT /Tx /T (HiddenField) /V (HIDDEN-FIELD-VALUE) /Kids [7 0 R] >>\nendobj",
            "7 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [60 20 180 40] /P 3 0 R /F 2 >>\nendobj",
            "8 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        Assert.Contains(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.Contains("HIDDEN-ANNOTATION", StringComparison.Ordinal));
        Assert.Contains(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.Contains("HIDDEN-FIELD-VALUE", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetyDoesNotClassifyValueAsHiddenWhenAnotherWidgetIsVisible() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 9 0 R >> >> /Contents 4 0 R /Annots [7 0 R 8 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "6 0 obj\n<< /FT /Tx /T (SharedField) /V (VISIBLE-SHARED-VALUE) /Kids [7 0 R 8 0 R] >>\nendobj",
            "7 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [20 20 100 40] /P 3 0 R /F 2 >>\nendobj",
            "8 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [120 20 220 40] /P 3 0 R /F 4 >>\nendobj",
            "9 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 10 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        Assert.DoesNotContain(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.Contains("VISIBLE-SHARED-VALUE", StringComparison.Ordinal));
    }

    private static byte[] BuildHiddenOptionalContentPdf() {
        const string pageContent = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET\n/OC /Hidden BDC BT /F1 12 Tf 20 100 Td (HIDDEN-PAGE-CONTENT) Tj ET EMC\n/OuterForm Do";
        const string outerFormContent = "/HiddenForm Do";
        const string formContent = "BT /F1 12 Tf 20 60 Td (HIDDEN-FORM-CONTENT) Tj ET";
        string pageStream = StreamObject(4, string.Empty, pageContent);
        string outerFormStream = StreamObject(
            7,
            "/Type /XObject /Subtype /Form /BBox [0 0 240 180] /Resources << /XObject << /HiddenForm 8 0 R >> >>",
            outerFormContent);
        string hiddenFormStream = StreamObject(
            8,
            "/Type /XObject /Subtype /Form /BBox [0 0 240 180] /OC 6 0 R /Resources << /Font << /F1 5 0 R >> >>",
            formContent);
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [6 0 R] /D << /BaseState /ON /OFF [6 0 R] >> >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> /Properties << /Hidden 6 0 R >> /XObject << /OuterForm 7 0 R >> >> /Contents 4 0 R >>\nendobj",
            pageStream,
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "6 0 obj\n<< /Type /OCG /Name (Hidden layer) >>\nendobj",
            outerFormStream,
            hiddenFormStream,
            "trailer\n<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static string StreamObject(int objectNumber, string dictionaryEntries, string content) {
        int length = Encoding.ASCII.GetByteCount(content);
        string entries = string.IsNullOrWhiteSpace(dictionaryEntries) ? string.Empty : dictionaryEntries + " ";
        return objectNumber + " 0 obj\n<< " + entries + "/Length " + length + " >>\nstream\n" + content + "\nendstream\nendobj";
    }
}
