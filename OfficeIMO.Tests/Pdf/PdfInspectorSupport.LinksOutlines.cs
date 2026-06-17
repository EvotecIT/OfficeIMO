using System;
using System.IO;
using System.Linq;
using System.Text;
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

    private static byte[] BuildWideOutlinePdf(int outlineItemCount) {
        var pdf = new StringBuilder();
        AppendMinimalOutlineDocumentHeader(pdf, outlineItemCount);

        for (int i = 0; i < outlineItemCount; i++) {
            int objectNumber = 6 + i;
            pdf.AppendLine(objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj");
            pdf.Append("<< /Title (Item ")
                .Append(i.ToString(System.Globalization.CultureInfo.InvariantCulture))
                .Append(") /Parent 5 0 R /Dest [3 0 R /Fit]");
            if (i + 1 < outlineItemCount) {
                pdf.Append(" /Next ")
                    .Append((objectNumber + 1).ToString(System.Globalization.CultureInfo.InvariantCulture))
                    .Append(" 0 R");
            }

            pdf.AppendLine(" >>");
            pdf.AppendLine("endobj");
        }

        pdf.AppendLine("trailer");
        pdf.Append("<< /Root 1 0 R /Size ")
            .Append((6 + outlineItemCount).ToString(System.Globalization.CultureInfo.InvariantCulture))
            .AppendLine(" >>");
        pdf.AppendLine("%%EOF");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }

    private static byte[] BuildDeepOutlinePdf(int outlineDepth) {
        var pdf = new StringBuilder();
        AppendMinimalOutlineDocumentHeader(pdf, 1);

        for (int i = 0; i < outlineDepth; i++) {
            int objectNumber = 6 + i;
            pdf.AppendLine(objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj");
            pdf.Append("<< /Title (Depth ")
                .Append((i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture))
                .Append(") /Parent ")
                .Append(i == 0 ? "5" : (objectNumber - 1).ToString(System.Globalization.CultureInfo.InvariantCulture))
                .Append(" 0 R /Dest [3 0 R /Fit]");
            if (i + 1 < outlineDepth) {
                pdf.Append(" /First ")
                    .Append((objectNumber + 1).ToString(System.Globalization.CultureInfo.InvariantCulture))
                    .Append(" 0 R");
            }

            pdf.AppendLine(" >>");
            pdf.AppendLine("endobj");
        }

        pdf.AppendLine("trailer");
        pdf.Append("<< /Root 1 0 R /Size ")
            .Append((6 + outlineDepth).ToString(System.Globalization.CultureInfo.InvariantCulture))
            .AppendLine(" >>");
        pdf.AppendLine("%%EOF");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }

    private static void AppendMinimalOutlineDocumentHeader(StringBuilder pdf, int outlineCount) {
        pdf.AppendLine("%PDF-1.4");
        pdf.AppendLine("1 0 obj");
        pdf.AppendLine("<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>");
        pdf.AppendLine("endobj");
        pdf.AppendLine("2 0 obj");
        pdf.AppendLine("<< /Type /Pages /Count 1 /Kids [3 0 R] >>");
        pdf.AppendLine("endobj");
        pdf.AppendLine("3 0 obj");
        pdf.AppendLine("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>");
        pdf.AppendLine("endobj");
        pdf.AppendLine("4 0 obj");
        pdf.AppendLine("<< /Length 0 >>");
        pdf.AppendLine("stream");
        pdf.AppendLine();
        pdf.AppendLine("endstream");
        pdf.AppendLine("endobj");
        pdf.AppendLine("5 0 obj");
        pdf.Append("<< /Type /Outlines /First 6 0 R /Last ")
            .Append((5 + outlineCount).ToString(System.Globalization.CultureInfo.InvariantCulture))
            .Append(" 0 R /Count ")
            .Append(outlineCount.ToString(System.Globalization.CultureInfo.InvariantCulture))
            .AppendLine(" >>");
        pdf.AppendLine("endobj");
    }


}
