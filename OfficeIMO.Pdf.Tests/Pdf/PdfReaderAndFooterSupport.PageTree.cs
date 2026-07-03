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

    private static byte[] BuildPdfWithContentStreamArray() {
        const string streamOne = "BT\n/F1 12 Tf\n72 720 Td\n(Hello) Tj\nET";
        const string streamTwo = "\nBT\n/F1 12 Tf\n72 720 Td\n( world) Tj\nET";
        int streamOneLength = Encoding.ASCII.GetByteCount(streamOne);
        int streamTwoLength = Encoding.ASCII.GetByteCount(streamTwo);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Contents [5 0 R 6 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamOneLength} >>",
            "stream",
            streamOne.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Length {streamTwoLength} >>",
            "stream",
            streamTwo.TrimEnd('\n'),
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithSplitTextStateContentStreamArray() {
        const string streamOne = "BT\n/F1 12 Tf\n72 720 Td";
        const string streamTwo = "\n(Split state) Tj\nET";
        int streamOneLength = Encoding.ASCII.GetByteCount(streamOne);
        int streamTwoLength = Encoding.ASCII.GetByteCount(streamTwo);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Contents [5 0 R 6 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamOneLength} >>",
            "stream",
            streamOne.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Length {streamTwoLength} >>",
            "stream",
            streamTwo.TrimEnd('\n'),
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithIndirectKidsArrayObject() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hello indirect kids) Tj\nET\n";
        int streamLength = Encoding.ASCII.GetByteCount(streamContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids 6 0 R /MediaBox [0 0 612 792] >>",
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
            "6 0 obj",
            "[3 0 R]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithCyclicKidsReferences() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hello cyclic kids) Tj\nET\n";
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
            "<< /Type /Pages /Parent 2 0 R /Kids [2 0 R 4 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Page /Parent 3 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 6 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "6 0 obj",
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

    private static byte[] BuildPdfWithIndirectContentArrayObject() {
        const string streamOne = "BT\n/F1 12 Tf\n72 720 Td\n(Hello) Tj\nET";
        const string streamTwo = "\nBT\n/F1 12 Tf\n72 720 Td\n( world) Tj\nET";
        int streamOneLength = Encoding.ASCII.GetByteCount(streamOne);
        int streamTwoLength = Encoding.ASCII.GetByteCount(streamTwo);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Contents 7 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamOneLength} >>",
            "stream",
            streamOne.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Length {streamTwoLength} >>",
            "stream",
            streamTwo.TrimEnd('\n'),
            "endstream",
            "endobj",
            "7 0 obj",
            "[5 0 R 6 0 R]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithCommentedPageDictionary() {
        const string streamContent = "BT\n/F1 12 Tf\n72 720 Td\n(Hello comments) Tj\nET\n";
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
            "<< /Type /Page",
            "/Parent 2 0 R",
            "/Resources % resources comment",
            "<< /Font << /F1 4 0 R >> >>",
            "/Contents % contents comment",
            "5 0 R",
            ">>",
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


    private static byte[] BuildPdfWithTwoDirectContentPages() {
        const string pageOne = "BT\n/F1 12 Tf\n72 720 Td\n(Direct page one) Tj\nET\n";
        const string pageTwo = "BT\n/F1 12 Tf\n72 720 Td\n(Direct page two) Tj\nET\n";

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 612 792] /Resources 5 0 R >>",
            "endobj",
            "3 0 obj",
            $"<< /Type /Page /Parent 2 0 R /Contents << /Length {Encoding.ASCII.GetByteCount(pageOne)} >>\nstream\n{pageOne.TrimEnd('\n')}\nendstream >>",
            "endobj",
            "4 0 obj",
            $"<< /Type /Page /Parent 2 0 R /Contents << /Length {Encoding.ASCII.GetByteCount(pageTwo)} >>\nstream\n{pageTwo.TrimEnd('\n')}\nendstream >>",
            "endobj",
            "5 0 obj",
            "<< /Font 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /F1 7 0 R >>",
            "endobj",
            "7 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithDistinctReferencedContentArrays() {
        const string sharedStream = "BT\n/F1 12 Tf\n72 720 Td\n(Shared stream page) Tj\nET\n";
        int sharedLength = Encoding.ASCII.GetByteCount(sharedStream);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 612 792] /Resources 8 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>",
            "endobj",
            "5 0 obj",
            "[7 0 R]",
            "endobj",
            "6 0 obj",
            "[7 0 R]",
            "endobj",
            "7 0 obj",
            $"<< /Length {sharedLength} >>",
            "stream",
            sharedStream.TrimEnd('\n'),
            "endstream",
            "endobj",
            "8 0 obj",
            "<< /Font 9 0 R >>",
            "endobj",
            "9 0 obj",
            "<< /F1 10 0 R >>",
            "endobj",
            "10 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

}
