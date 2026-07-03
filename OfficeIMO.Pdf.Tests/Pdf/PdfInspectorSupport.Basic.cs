using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    private static byte[] BuildTwoPagePdf() {
        var doc = PdfDocument.Create()
            .Meta(
                title: "Inspection sample",
                author: "OfficeIMO",
                subject: "Roadmap",
                keywords: "pdf,inspect");

        doc.Compose(compose => {
            compose.Page(page => {
                page.Size(PageSizes.A4);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("First page."))));
            });

            compose.Page(page => {
                page.Size(new PageSize(792, 612));
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Landscape page."))));
            });
        });

        return doc.ToBytes();
    }

    private static byte[] BuildUnsupportedContentStreamFilterPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 4 /Filter /DCTDecode >>",
            "stream",
            "data",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildUnsupportedFormXObjectStreamFilterPdf() {
        const string pageContent = "q\n/Fm1 Do\nQ";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /XObject << /Fm1 5 0 R >> >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + pageContent.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            pageContent,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Form /BBox [0 0 200 200] /Length 4 /Filter /DCTDecode >>",
            "stream",
            "data",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildUnsupportedFormXObjectFilterSplitAcrossContentStreamsPdf(string pageContentOne = "q\n/Fm1 ", string pageContentTwo = "Do\nQ") {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /XObject << /Fm1 5 0 R >> >> /Contents [4 0 R 6 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + pageContentOne.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            pageContentOne,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Form /BBox [0 0 200 200] /Length 4 /Filter /DCTDecode >>",
            "stream",
            "data",
            "endstream",
            "endobj",
            "6 0 obj",
            "<< /Length " + pageContentTwo.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            pageContentTwo,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildWrongGenerationContentReferencePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 1 R >>",
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

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

}
