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

    private static byte[] BuildSingleStreamPdf(string streamContent) {
        return BuildSingleStreamPdf(Encoding.ASCII.GetBytes(streamContent.TrimEnd('\n')));
    }

    private static byte[] BuildSingleStreamPdfWithMarkedContentProperties(string streamContent, string actualTextPdfString = "<FEFF005200650073006F00750072006300650020005A00650064>") {
        streamContent = streamContent.TrimEnd('\n');
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
            $"<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> /Properties << /MC0 << /ActualText {actualTextPdfString} >> >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamLength} >>",
            "stream",
            streamContent,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithFontEncodingDifferences(
        string differences = "65 /Z /space /Euro /uni0104 /A.alt",
        string encodedText = "4142434445") {
        string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n<" + encodedText + "> Tj\nET\n";
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
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /BaseEncoding /WinAnsiEncoding /Differences [" + differences + "] >> >>",
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

    private static byte[] BuildSingleStreamPdf(byte[] streamBytes, string extraStreamDictionaryEntries = "") {
        return BuildSingleStreamPdfWithExtraObjects(streamBytes, extraStreamDictionaryEntries);
    }

    private static byte[] BuildSingleStreamPdfWithExtraObjects(byte[] streamBytes, string extraStreamDictionaryEntries = "", params string[] extraObjects) {
        using var ms = new MemoryStream();
        using var writer = new StreamWriter(ms, Encoding.ASCII, 1024, leaveOpen: true);

        writer.WriteLine("%PDF-1.4");
        writer.WriteLine("1 0 obj");
        writer.WriteLine("<< /Type /Catalog /Pages 2 0 R >>");
        writer.WriteLine("endobj");
        writer.WriteLine("2 0 obj");
        writer.WriteLine("<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>");
        writer.WriteLine("endobj");
        writer.WriteLine("3 0 obj");
        writer.WriteLine("<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>");
        writer.WriteLine("endobj");
        writer.WriteLine("4 0 obj");
        writer.WriteLine("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
        writer.WriteLine("endobj");
        writer.WriteLine("5 0 obj");
        writer.Write("<< /Length ");
        writer.Write(streamBytes.Length);
        if (!string.IsNullOrWhiteSpace(extraStreamDictionaryEntries)) {
            writer.Write(' ');
            writer.Write(extraStreamDictionaryEntries.Trim());
        }
        writer.WriteLine(" >>");
        writer.WriteLine("stream");
        writer.Flush();

        ms.Write(streamBytes, 0, streamBytes.Length);

        writer.WriteLine();
        writer.WriteLine("endstream");
        writer.WriteLine("endobj");
        foreach (string extraObject in extraObjects) {
            writer.WriteLine(extraObject);
        }
        writer.WriteLine("trailer");
        writer.WriteLine("<< /Root 1 0 R >>");
        writer.WriteLine("%%EOF");
        writer.Flush();
        return ms.ToArray();
    }

}
