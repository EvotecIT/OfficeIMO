using System;
using System.IO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReadStreamTests {
    private static byte[] BuildPdf() {
        return PdfDocument.Create()
            .Meta(title: "Stream read", author: "OfficeIMO", subject: "Read", keywords: "stream,pdf")
            .Paragraph(p => p.Text("Stream readable text"))
            .ToBytes();
    }

    private static byte[] BuildTwoPageLinkAnnotationPdf() {
        return PdfDocument.Create()
            .Paragraph(p => p.Link("First", "https://evotec.xyz/first", contents: "First link metadata"))
            .PageBreak()
            .Paragraph(p => p.Link("Second", "https://evotec.xyz/second", contents: "Second link metadata"))
            .ToBytes();
    }

    private static byte[] BuildGoToActionLinkAnnotationPdf(string destination = "[3 0 R /FitH 144]") {
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
            "<< /Type /Annot /Subtype /Link /Rect [10 20 90 42] /Contents (Jump to top) /A << /S /GoTo /D " + destination + " >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildOutlinePdf() {
        return PdfDocument.Create(new PdfOptions { CreateOutlineFromHeadings = true })
            .H1("Executive summary")
            .Paragraph(p => p.Text("Outline sample"))
            .ToBytes();
    }

    private static byte[] BuildTwoPageOutlinePdf() {
        return PdfDocument.Create(new PdfOptions { CreateOutlineFromHeadings = true })
            .H1("Executive summary")
            .Paragraph(p => p.Text("Outline sample"))
            .PageBreak()
            .H1("Appendix")
            .Paragraph(p => p.Text("Appendix sample"))
            .ToBytes();
    }

    private static byte[] BuildGoToActionOutlinePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>",
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
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Chapter 1) /Parent 5 0 R /A << /S /GoTo /D [3 0 R /XYZ 0 200 0] >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFitHorizontalOutlinePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>",
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
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Fit horizontal) /Parent 5 0 R /Dest [3 0 R /FitH 144] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildGoToActionIndirectDestinationPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>",
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
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Indirect GoTo action) /Parent 5 0 R /A << /S /GoTo /D 7 0 R >> >>",
            "endobj",
            "7 0 obj",
            "[3 0 R /XYZ 0 188 0]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildGoToActionDictionaryDestinationPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>",
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
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Dictionary GoTo action) /Parent 5 0 R /A << /S /GoTo /D << /D [3 0 R /XYZ 0 188 0] >> >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCyclicGoToActionDestinationPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>",
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
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Cyclic GoTo action) /Parent 5 0 R /A << /S /GoTo /D 7 0 R >> >>",
            "endobj",
            "7 0 obj",
            "8 0 R",
            "endobj",
            "8 0 obj",
            "7 0 R",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildUriActionOutlinePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>",
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
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (External) /Parent 5 0 R /A << /S /URI /URI (https://evotec.xyz) >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildIndirectOutlineDestinationPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>",
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
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Indirect destination) /Parent 5 0 R /Dest 7 0 R >>",
            "endobj",
            "7 0 obj",
            "[3 0 R /XYZ 0 144 0]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildDirectNamedDestinationOutlinePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines /Dests 7 0 R >>",
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
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Direct named destination) /Parent 5 0 R /Dest /Chapter1 >>",
            "endobj",
            "7 0 obj",
            "<< /Chapter1 [3 0 R /XYZ 0 200 0] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildNameTreeNamedDestinationOutlinePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines /Names << /Dests << /Names [(Chapter1) [3 0 R /XYZ 0 188 0]] >> >> >>",
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
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Name-tree named destination) /Parent 5 0 R /Dest (Chapter1) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildGoToActionNamedDestinationOutlinePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines /Dests 7 0 R >>",
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
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Action named destination) /Parent 5 0 R /A << /S /GoTo /D /Chapter1 >> >>",
            "endobj",
            "7 0 obj",
            "<< /Chapter1 [3 0 R /XYZ 0 176 0] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildStaleCatalogRevisionPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageLayout /TwoColumnLeft /PageLabels 6 0 R >>",
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
            "<< /Type /Catalog /Pages 2 0 R /PageLayout /SinglePage >>",
            "endobj",
            "6 0 obj",
            "<< /Kids [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /Nums [0 << /S /D /St 10 >>] >>",
            "endobj",
            "trailer",
            "<< /Root 5 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildMixedNamedDestinationOutlinePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines /Dests 9 0 R /Names << /Dests << /Names [(Chapter1) [3 0 R /XYZ 0 188 0]] >> >> >>",
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
            "<< /Type /Outlines /First 6 0 R /Last 8 0 R /Count 3 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Direct name destination) /Parent 5 0 R /Dest /Chapter1 /Next 7 0 R >>",
            "endobj",
            "7 0 obj",
            "<< /Title (Name-tree string destination) /Parent 5 0 R /Dest (Chapter1) /Prev 6 0 R /Next 8 0 R >>",
            "endobj",
            "8 0 obj",
            "<< /Title (Dictionary string destination) /Parent 5 0 R /Dest << /D (Chapter1) >> /Prev 7 0 R >>",
            "endobj",
            "9 0 obj",
            "<< /Chapter1 [3 0 R /XYZ 0 144 0] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 10 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildStaleCatalogWithDifferentPagesAndOutlinesPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 4 0 R /PageLayout /TwoColumnLeft >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 11 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Outlines /First 12 0 R /Last 12 0 R /Count 1 >>",
            "endobj",
            "5 0 obj",
            "<< /Type /Catalog /Pages 6 0 R /Outlines 8 0 R /PageLayout /SinglePage >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Pages /Count 1 /Kids [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /Type /Page /Parent 6 0 R /MediaBox [0 0 200 200] /Contents 11 0 R >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Outlines /First 9 0 R /Last 9 0 R /Count 1 >>",
            "endobj",
            "9 0 obj",
            "<< /Title (Current) /Parent 8 0 R /Dest [7 0 R /XYZ 0 144 0] >>",
            "endobj",
            "11 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "12 0 obj",
            "<< /Title (Old) /Parent 4 0 R /Dest [3 0 R /XYZ 0 72 0] >>",
            "endobj",
            "trailer",
            "<< /Root 5 0 R /Size 13 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }


}
