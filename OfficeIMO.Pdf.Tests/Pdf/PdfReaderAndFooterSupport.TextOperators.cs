using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReaderAndFooterRegressionTests {

    private static byte[] BuildPdfWithTjArraySpacing() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n[(Hello) -600 (world)] TJ\nET\n";
        int streamLength = Encoding.ASCII.GetByteCount(streamContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamLength} >>",
            "stream",
            streamContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static void AssertSecondLineStartsAtFirstLineX(byte[] bytes) {
        var spans = PdfReadDocument.Open(bytes).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan first = Assert.Single(spans, span => span.Text == "First");
        PdfTextSpan second = Assert.Single(spans, span => span.Text == "Second");

        Assert.Equal(first.X, second.X, 2);
        Assert.True(second.Y < first.Y, $"Expected second line Y {second.Y} to be below first line Y {first.Y}.");
    }

    private static byte[] BuildPdfWithSingleQuoteOperator() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hello) Tj\n( world) '\nET\n";
        return BuildSingleStreamPdf(streamContent);
    }

    private static byte[] BuildPdfWithDoubleQuoteOperator() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hello) Tj\n0 0 ( world) \"\nET\n";
        return BuildSingleStreamPdf(streamContent);
    }

    private static byte[] BuildPdfWithDoubleQuoteLineAdvanceOperator() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(First) Tj\n0 0 (Second) \"\nET\n";
        return BuildSingleStreamPdf(streamContent);
    }

    private static byte[] BuildPdfWithTDTextPositioning() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(First) Tj\n0 -14 TD\n(Second) Tj\nET\n";
        return BuildSingleStreamPdf(streamContent);
    }

    private static byte[] BuildPdfWithInitialTdTextPositioning() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(First) Tj\nET\nBT\n/F1 12 Tf\n110 720 Td\n(Second) Tj\nET\n";
        return BuildSingleStreamPdf(streamContent);
    }

    private static byte[] BuildPdfWithRepeatedTDTextPositioning() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(First) Tj\n0 -14 TD\n0 -14 TD\n(Second) Tj\nET\n";
        return BuildSingleStreamPdf(streamContent);
    }


    private static byte[] BuildPdfWithHexMetadata(string title, string author) {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 0 /Kids [ ] >>",
            "endobj",
            "3 0 obj",
            $"<< /Title <{EncodeUtf16BeHex(title)}> /Author <{EncodeUtf16BeHex(author)}> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Info 3 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithLiteralMetadata(string titleLiteral, string authorLiteral) {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 0 /Kids [ ] >>",
            "endobj",
            "3 0 obj",
            $"<< /Title {titleLiteral} /Author {authorLiteral} >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Info 3 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }


    private static string EncodeUtf16BeHex(string value) {
        byte[] bom = new byte[] { 0xFE, 0xFF };
        byte[] textBytes = Encoding.BigEndianUnicode.GetBytes(value);
        byte[] bytes = new byte[bom.Length + textBytes.Length];
        Buffer.BlockCopy(bom, 0, bytes, 0, bom.Length);
        Buffer.BlockCopy(textBytes, 0, bytes, bom.Length, textBytes.Length);

        var sb = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            sb.Append(bytes[i].ToString("X2"));
        }
        return sb.ToString();
    }
}
