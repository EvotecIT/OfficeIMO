using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfLogicalDocumentTests {
    private static string Normalize(string text) {
        return new string(text.Where(ch => !char.IsWhiteSpace(ch)).ToArray());
    }

    private static bool RowContains(IReadOnlyList<string> row, params string[] expectedTokens) {
        string rowText = Normalize(string.Join(" ", row));
        return expectedTokens.All(token => rowText.Contains(token, StringComparison.Ordinal));
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int index = 0;
        while (true) {
            index = text.IndexOf(value, index, StringComparison.Ordinal);
            if (index < 0) {
                return count;
            }

            count++;
            index += value.Length;
        }
    }

    private static void AssertContainsInOrder(string text, params string[] expectedTokens) {
        int lastIndex = -1;
        for (int i = 0; i < expectedTokens.Length; i++) {
            int index = text.IndexOf(expectedTokens[i], StringComparison.Ordinal);
            Assert.True(index >= 0, $"Expected token '{expectedTokens[i]}' was not found.");
            Assert.True(index > lastIndex, $"Expected token '{expectedTokens[i]}' to appear after the previous token.");
            lastIndex = index;
        }
    }

    private static byte[] BuildThreePageLogicalPdf() {
        return PdfDocument.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 180,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("First logical page marker."))
            .PageBreak()
            .Paragraph(p => p.Text("Second logical page marker."))
            .PageBreak()
            .Paragraph(p => p.Text("Third logical page marker."))
            .ToBytes();
    }

    private static byte[] BuildHierarchicalFormPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
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
            "<< /NeedAppearances true /SigFlags 1 /DA (/Helv 7 Tf 0.5 g) /Fields [6 0 R 8 0 R 9 0 R] >>",
            "endobj",
            "6 0 obj",
            "<< /FT /Tx /T (Person) /Ff 1 /MaxLen 64 /DV (InheritedDraft) /DA (/Helv 10 Tf 0 g) /Q 2 /Kids [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /T (Name) /TU (Display name) /TM (ExportName) /V (OfficeIMO) >>",
            "endobj",
            "8 0 obj",
            "<< /FT /Btn /T (AcceptTerms) /V /Yes >>",
            "endobj",
            "9 0 obj",
            "<< /FT /Ch /T (Selection) /V /DE /DV (PL) /DA (/Helv 8 Tf 0 0 1 rg) /Q 1 /Opt [[(PL) (Poland)] (DE)] /Kids [10 0 R] >>",
            "endobj",
            "10 0 obj",
            "<< /T (Country) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 11 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFieldKindFormPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
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
            "<< /Fields [6 0 R 7 0 R 8 0 R 9 0 R 10 0 R 11 0 R] >>",
            "endobj",
            "6 0 obj",
            "<< /FT /Tx /T (Notes) /V (Secret) /DV (Draft) /Ff 12295 /MaxLen 42 /DA (/Helv 9 Tf 0 g) /Q 0 >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Btn /T (Accept) /V /Yes >>",
            "endobj",
            "8 0 obj",
            "<< /FT /Btn /T (Choice) /V /A /Ff 49152 >>",
            "endobj",
            "9 0 obj",
            "<< /FT /Btn /T (Submit) /Ff 65536 >>",
            "endobj",
            "10 0 obj",
            "<< /FT /Ch /T (Country) /V [(PL) /US] /DV [(DE) /US] /Ff 74317826 /DA (/Helv 8 Tf 0 0 1 rg) /Q 1 /Opt [[(PL) (Poland)] (DE) [/US (United States)]] >>",
            "endobj",
            "11 0 obj",
            "<< /FT /Sig /T (Approval) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 12 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildWidgetFormPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [8 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Btn /T (AcceptTerms) /V /Yes /Kids [8 0 R] >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 100 36 116] /F 4 /AS /Yes /AP << /N << /Off 9 0 R /Yes 10 0 R >> >> >>",
            "endobj",
            "9 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "10 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 11 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildNavigationPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageMode /FullScreen /PageLayout /TwoColumnLeft /Version /1.7 /Lang (en-US) /PageLabels 5 0 R /Dests 6 0 R /OpenAction [3 0 R /XYZ 0 200 0] /ViewerPreferences 7 0 R /Outlines 8 0 R >>",
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
            "<< /Nums [0 << /S /D /P (A-) /St 3 >>] >>",
            "endobj",
            "6 0 obj",
            "<< /Chapter1 [3 0 R /XYZ 0 200 0] >>",
            "endobj",
            "7 0 obj",
            "<< /HideToolbar true /DisplayDocTitle true >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Outlines /First 9 0 R /Last 9 0 R /Count 1 >>",
            "endobj",
            "9 0 obj",
            "<< /Title (Logical outline) /Parent 8 0 R /Dest [3 0 R /XYZ 0 200 0] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 10 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildDirectDestinationLinkPdf(string destination = "[3 0 R /FitR 10 20 90 144]") {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
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
            "<< /Type /Annot /Subtype /Link /Rect [10 20 90 42] /Contents (Direct destination link) /A << /S /GoTo /D " + destination + " >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildNamedActionLinkPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
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
            "<< /Type /Annot /Subtype /Link /Rect [10 20 90 42] /Contents (Next page action) /A << /S /Named /N /NextPage >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCatalogJavaScriptActionPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Kids [6 0 R] >> >> /OpenAction 7 0 R /AA << /WC 8 0 R >> >>",
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
            "<< /S /JavaScript /JS (app.alert('OfficeIMO')) >>",
            "endobj",
            "6 0 obj",
            "<< /Names [(Open) 5 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /S /JavaScript /JS (app.alert('OpenAction')) >>",
            "endobj",
            "8 0 obj",
            "<< /S /Launch /F (tool.exe) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPageAdditionalActionsPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /AA << /O << /S /JavaScript /JS (app.alert('Page open')) >> /C << /S /Launch /F (tool.exe) >> >> >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPageChainedActionsPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /AA << /O << /S /JavaScript /JS (app.alert('Page open')) /Next [5 0 R 6 0 R] >> >> >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /S /Launch /F (tool.exe) >>",
            "endobj",
            "6 0 obj",
            "<< /S /RichMedia >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildRemoteGoToLinkPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
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
            "<< /Type /Annot /Subtype /Link /Rect [10 20 90 42] /Contents (Remote report link) /A << /S /GoToR /F << /F (fallback.pdf) /UF (remote-report.pdf) >> /D [1 /FitH 144] >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildThreePageNavigationPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Dests 9 0 R /OpenAction [7 0 R /Fit] /Outlines 10 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 3 /Kids [3 0 R 5 0 R 7 0 R] >>",
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
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 8 0 R >>",
            "endobj",
            "8 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "9 0 obj",
            "<< /First [3 0 R /XYZ 0 200 0] /Second [5 0 R /XYZ 0 200 0] /Third [7 0 R /XYZ 0 200 0] >>",
            "endobj",
            "10 0 obj",
            "<< /Type /Outlines /First 11 0 R /Last 13 0 R /Count 3 >>",
            "endobj",
            "11 0 obj",
            "<< /Title (First outline) /Parent 10 0 R /Next 12 0 R /Dest [3 0 R /XYZ 0 200 0] >>",
            "endobj",
            "12 0 obj",
            "<< /Title (Second outline) /Parent 10 0 R /Prev 11 0 R /Next 13 0 R /Dest [5 0 R /FitR 10 20 90 144] >>",
            "endobj",
            "13 0 obj",
            "<< /Title (Third outline) /Parent 10 0 R /Prev 12 0 R /Dest [7 0 R /XYZ 0 200 0] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 14 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildThreePageLabelPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageLabels 9 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 3 /Kids [3 0 R 5 0 R 7 0 R] >>",
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
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 8 0 R >>",
            "endobj",
            "8 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "9 0 obj",
            "<< /Nums [0 << /S /D /P (A-) /St 10 >> 2 << /S /r /P (B-) /St 3 >>] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 10 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreateMinimalRgbPng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 12,
            73, 68, 65, 84,
            0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }
}
