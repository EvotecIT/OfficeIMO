using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    private static byte[] BuildAnnotatedPdf() {
        return PdfDocument.Create()
            .Paragraph(p => p.Link("OfficeIMO link", "https://evotec.xyz"))
            .ToBytes();
    }

    private static byte[] BuildOutlinePdf() {
        return PdfDocument.Create(new PdfOptions { CreateOutlineFromHeadings = true })
            .H1("Executive summary")
            .Paragraph(p => p.Text("Outline sample"))
            .ToBytes();
    }

    private static byte[] BuildComplexOutlinePdf() {
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


}
