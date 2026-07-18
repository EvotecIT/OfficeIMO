using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfParsingModeTests {
    [Fact]
    public void LenientModeReportsAndRecoversIncorrectStreamLength() {
        byte[] pdf = BuildStreamPdf("/Length 999");

        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Lenient });

        PdfRepairDiagnostic repair = Assert.Single(document.RepairReport.Diagnostics, item => item.Code == "IncorrectStreamLength");
        Assert.Equal(4, repair.ObjectNumber);
        Assert.Contains("declares /Length 999", repair.Message, StringComparison.Ordinal);
        Assert.Contains("Recovered stream text", document.ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void StrictModeRejectsIncorrectStreamLengthWithStableCode() {
        byte[] pdf = BuildStreamPdf("/Length 999");

        PdfParseException exception = Assert.Throws<PdfParseException>(() =>
            PdfReadDocument.Open(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Strict }));

        Assert.Equal("IncorrectStreamLength", exception.Code);
        Assert.Equal(4, exception.ObjectNumber);
    }

    [Fact]
    public void MissingStreamLengthIsReportedOrRejectedByPolicy() {
        byte[] pdf = BuildStreamPdf(string.Empty);

        PdfReadDocument lenient = PdfReadDocument.Open(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Lenient });
        PdfParseException strict = Assert.Throws<PdfParseException>(() =>
            PdfReadDocument.Open(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Strict }));

        Assert.Contains(lenient.RepairReport.Diagnostics, item => item.Code == "MissingStreamLength" && item.ObjectNumber == 4);
        Assert.Equal("MissingStreamLength", strict.Code);
    }

    [Fact]
    public void BinaryEndStreamBytesDoNotCreateFalseMissingEndObjectRepair() {
        const string streamData = "prefix endstream\n5 0 obj\nbinary-like suffix";
        byte[] pdf = BuildStreamPdf("/Length " + streamData.Length, streamData);

        PdfReadDocument document = PdfReadDocument.Open(pdf);
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfStream stream = Assert.IsType<PdfStream>(objects[4].Value);

        Assert.Equal(streamData, Encoding.ASCII.GetString(stream.Data));
        Assert.False(objects.ContainsKey(5));
        Assert.DoesNotContain(
            document.RepairReport.Diagnostics,
            diagnostic => diagnostic.Code == "MissingEndObject");
    }

    [Fact]
    public void IndirectStreamLength_BoundsObjectLikeBytesInsideValidStream() {
        const string streamData = "prefix endstream\n5 0 obj\nbinary-like suffix";
        byte[] pdf = BuildStreamPdf(
            "/Length 6 0 R",
            streamData,
            "6 0 obj\n" + streamData.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\nendobj");

        var (objects, _) = PdfSyntax.ParseObjects(
            pdf,
            new PdfReadOptions { ParsingMode = PdfParsingMode.Lenient },
            out PdfRepairReport repairReport);

        PdfStream stream = Assert.IsType<PdfStream>(objects[4].Value);
        Assert.Equal(streamData, Encoding.ASCII.GetString(stream.Data));
        Assert.False(objects.ContainsKey(5));
        Assert.IsType<PdfNumber>(objects[6].Value);
        Assert.DoesNotContain(
            repairReport.Diagnostics,
            diagnostic => diagnostic.Code == "MissingEndObject" ||
                diagnostic.Code == "IncorrectStreamLength");
    }

    [Fact]
    public void IndirectStreamLength_IgnoresMatchingScalarBytesInsideLaterStream() {
        const string firstStreamData = "prefix endstream\n5 0 obj\nbinary-like suffix";
        const string laterStreamData = "6 0 obj\n1\nendobj\npayload";
        byte[] pdf = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n" +
            "4 0 obj\n<< /Length 6 0 R >>\nstream\n" +
            firstStreamData +
            "\nendstream\nendobj\n" +
            "6 0 obj\n" +
            firstStreamData.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            "\nendobj\n" +
            "7 0 obj\n<< /Length " +
            laterStreamData.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n" +
            laterStreamData +
            "\nendstream\nendobj\n" +
            "trailer\n<< /Root 1 0 R /Size 8 >>\nstartxref\n0\n%%EOF\n");

        var (objects, _) = PdfSyntax.ParseObjects(pdf);

        PdfStream firstStream = Assert.IsType<PdfStream>(objects[4].Value);
        PdfNumber length = Assert.IsType<PdfNumber>(objects[6].Value);
        PdfStream laterStream = Assert.IsType<PdfStream>(objects[7].Value);
        Assert.Equal(firstStreamData, Encoding.ASCII.GetString(firstStream.Data));
        Assert.Equal(firstStreamData.Length, length.Value);
        Assert.Equal(laterStreamData, Encoding.ASCII.GetString(laterStream.Data));
        Assert.False(objects.ContainsKey(5));
    }

    [Fact]
    public void IndirectStreamLength_ResolvesChainsBeyondFourStreamRanges() {
        const int indirectStreamCount = 6;
        const int firstStreamObject = 10;
        const int firstLengthObject = 100;
        var streamData = new string[indirectStreamCount + 1];
        streamData[0] = "root payload";
        for (int i = 1; i < streamData.Length; i++) {
            streamData[i] =
                (firstLengthObject + i - 1).ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " 0 obj\n1\nendobj\npayload-" +
                i.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        var pdf = new StringBuilder(
            "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n");
        for (int i = 0; i < indirectStreamCount; i++) {
            pdf.Append(firstLengthObject + i)
                .Append(" 0 obj\n")
                .Append(streamData[i].Length)
                .Append("\nendobj\n");
        }

        for (int i = 0; i < streamData.Length; i++) {
            pdf.Append(firstStreamObject + i).Append(" 0 obj\n<< /Length ");
            if (i < indirectStreamCount) {
                pdf.Append(firstLengthObject + i).Append(" 0 R");
            } else {
                pdf.Append(streamData[i].Length);
            }

            pdf.Append(" >>\nstream\n")
                .Append(streamData[i])
                .Append("\nendstream\nendobj\n");
        }

        pdf.Append("trailer\n<< /Root 1 0 R /Size 106 >>\nstartxref\n0\n%%EOF\n");

        var (objects, _) = PdfSyntax.ParseObjects(Encoding.ASCII.GetBytes(pdf.ToString()));

        for (int i = 0; i < streamData.Length; i++) {
            PdfStream stream = Assert.IsType<PdfStream>(objects[firstStreamObject + i].Value);
            Assert.Equal(streamData[i], Encoding.ASCII.GetString(stream.Data));
        }

        for (int i = 0; i < indirectStreamCount; i++) {
            PdfNumber length = Assert.IsType<PdfNumber>(objects[firstLengthObject + i].Value);
            Assert.Equal(streamData[i].Length, length.Value);
        }
    }

    [Fact]
    public void IndirectStreamLength_IgnoresLimitFailuresFromObjectLikePayloadBytes() {
        const string streamData = "5 0 obj\n<< /A [1 2 3 4 5 6 7 8 9 10] >>\nendobj\npayload";
        byte[] pdf = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "6 0 obj\n" +
            streamData.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            "\nendobj\n" +
            "4 0 obj\n<< /Length 6 0 R >>\nstream\n" +
            streamData +
            "\nendstream\nendobj\n" +
            "trailer\n<< /Size 7 >>\nstartxref\n0\n%%EOF\n");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxTokensPerObject = 8 }
        };

        var (objects, _) = PdfSyntax.ParseObjects(pdf, options);

        PdfStream stream = Assert.IsType<PdfStream>(objects[4].Value);
        Assert.Equal(streamData, Encoding.ASCII.GetString(stream.Data));
        Assert.False(objects.ContainsKey(5));
    }

    [Fact]
    public void MissingStreamBoundary_DoesNotConsumeFollowingValidStreamObject() {
        byte[] pdf = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n" +
            "4 0 obj\n<< /Length 5 >>\nstream\nabcde\n" +
            "5 0 obj\n<< /Length 2 >>\nstream\nOK\nendstream\nendobj\n" +
            "trailer\n<< /Root 1 0 R /Size 6 >>\nstartxref\n0\n%%EOF\n");

        var (objects, _) = PdfSyntax.ParseObjects(
            pdf,
            new PdfReadOptions { ParsingMode = PdfParsingMode.Lenient },
            out PdfRepairReport repairReport);

        PdfStream followingStream = Assert.IsType<PdfStream>(objects[5].Value);
        Assert.Equal("OK", Encoding.ASCII.GetString(followingStream.Data));
        Assert.Contains(
            repairReport.Diagnostics,
            diagnostic => diagnostic.Code == "MissingEndObject" && diagnostic.ObjectNumber == 4);
    }

    [Fact]
    public void MissingEndObjectIsReportedOrRejectedByPolicy() {
        byte[] pdf = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n" +
            "3 0 obj\n<< /Producer (unterminated object boundary) >>\n" +
            "trailer\n<< /Root 1 0 R /Size 4 >>\nstartxref\n0\n%%EOF\n");

        PdfReadDocument lenient = PdfReadDocument.Open(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Lenient });
        PdfParseException strict = Assert.Throws<PdfParseException>(() =>
            PdfReadDocument.Open(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Strict }));

        Assert.Contains(lenient.RepairReport.Diagnostics, item => item.Code == "MissingEndObject" && item.ObjectNumber == 3);
        Assert.Equal("MissingEndObject", strict.Code);
        Assert.Equal(3, strict.ObjectNumber);
    }

    [Fact]
    public void CleanDocumentHasEmptyRepairReport() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("Clean parse")).ToBytes();

        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Strict });

        Assert.False(document.RepairReport.HasRepairs);
        Assert.Empty(document.RepairReport.Diagnostics);
    }

    [Fact]
    public void InvalidStartXrefIsExplicitlyRebuiltOrRejectedByPolicy() {
        byte[] clean = PdfDocument.Create().Paragraph(p => p.Text("Cross-reference recovery")).ToBytes();
        byte[] damaged = ReplaceLastStartXrefOffset(clean, "1");

        PdfReadDocument lenient = PdfReadDocument.Open(damaged, new PdfReadOptions { ParsingMode = PdfParsingMode.Lenient });
        PdfParseException strict = Assert.Throws<PdfParseException>(() =>
            PdfReadDocument.Open(damaged, new PdfReadOptions { ParsingMode = PdfParsingMode.Strict }));

        PdfRepairDiagnostic repair = Assert.Single(lenient.RepairReport.Diagnostics, item => item.Code == "InvalidStartXref");
        Assert.Contains("rebuilt the object index", repair.Message, StringComparison.Ordinal);
        Assert.Contains("Cross-reference recovery", lenient.ExtractText(), StringComparison.Ordinal);
        Assert.Equal("InvalidStartXref", strict.Code);
    }

    [Fact]
    public void MissingStartXrefIsExplicitlyRebuiltOrRejectedByPolicy() {
        byte[] pdf = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n" +
            "trailer\n<< /Root 1 0 R /Size 3 >>\n%%EOF\n");

        PdfReadDocument lenient = PdfReadDocument.Open(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Lenient });
        PdfParseException strict = Assert.Throws<PdfParseException>(() =>
            PdfReadDocument.Open(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Strict }));

        Assert.Contains(lenient.RepairReport.Diagnostics, item => item.Code == "MissingStartXref");
        Assert.Equal("MissingStartXref", strict.Code);
    }

    [Fact]
    public void IncorrectPageCountDoesNotHideReachablePagesAndStrictModeRejectsIt() {
        byte[] clean = PdfDocument.Create().Paragraph(p => p.Text("Page one")).PageBreak().Paragraph(p => p.Text("Page two")).ToBytes();
        string text = Encoding.ASCII.GetString(clean);
        Assert.Contains("/Count 2", text, StringComparison.Ordinal);
        byte[] damaged = Encoding.ASCII.GetBytes(text.Replace("/Count 2", "/Count 1"));

        PdfReadDocument lenient = PdfReadDocument.Open(damaged, new PdfReadOptions { ParsingMode = PdfParsingMode.Lenient });
        PdfParseException strict = Assert.Throws<PdfParseException>(() => PdfReadDocument.Open(damaged, new PdfReadOptions { ParsingMode = PdfParsingMode.Strict }));

        Assert.Equal(2, lenient.Pages.Count);
        PdfRepairDiagnostic diagnostic = Assert.Single(lenient.RepairReport.Diagnostics, item => item.Code == "IncorrectPageTreeCount");
        Assert.True(diagnostic.WasRecovered);
        Assert.Equal("IncorrectPageTreeCount", strict.Code);
    }

    [Fact]
    public void MalformedNameTreeBrokenDestinationAndOrphanAreDiagnosedWithoutDestructiveRepair() {
        byte[] pdf = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /Names << /Dests 7 0 R >> >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 300 300] >>\nendobj\n" +
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>\nendobj\n" +
            "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n" +
            "7 0 obj\n<< /Names [(bad) [99 0 R /Fit] (dangling)] >>\nendobj\n" +
            "8 0 obj\n<< /Type /Page /MediaBox [0 0 10 10] >>\nendobj\n" +
            "trailer\n<< /Root 1 0 R /Size 9 >>\n%%EOF\n");

        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Lenient });

        Assert.Contains(document.RepairReport.Diagnostics, item => item.Code == "OddNameTreePairs" && !item.WasRecovered);
        Assert.Contains(document.RepairReport.Diagnostics, item => item.Code == "BrokenNamedDestination" && !item.WasRecovered);
        Assert.Contains(document.RepairReport.Diagnostics, item => item.Code == "OrphanedSemanticObjects" && item.ObjectNumber == 8 && !item.WasRecovered);
        Assert.True(document.RepairReport.DetectionOnlyCount >= 3);
    }

    [Fact]
    public void DuplicateObjectIdentifierWithoutIncrementalChainUsesLastDefinitionAndReportsRecovery() {
        byte[] pdf = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n" +
            "7 0 obj\n(first)\nendobj\n" +
            "7 0 obj\n(second)\nendobj\n" +
            "trailer\n<< /Root 1 0 R /Size 8 >>\n%%EOF\n");

        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Lenient });

        PdfRepairDiagnostic duplicate = Assert.Single(document.RepairReport.Diagnostics, item => item.Code == "DuplicateObjectIdentifier");
        Assert.Equal(7, duplicate.ObjectNumber);
        Assert.True(duplicate.WasRecovered);
    }

    private static byte[] BuildStreamPdf(
        string lengthEntry,
        string streamData = "BT (Recovered stream text) Tj ET",
        string extraObjects = "") {
        string dictionary = string.IsNullOrEmpty(lengthEntry) ? "<< >>" : "<< " + lengthEntry + " >>";
        string additionalObjects = string.IsNullOrEmpty(extraObjects)
            ? string.Empty
            : extraObjects.TrimEnd('\r', '\n') + "\n";
        return Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R >>\nendobj\n" +
            "4 0 obj\n" + dictionary + "\nstream\n" + streamData + "\nendstream\nendobj\n" +
            additionalObjects +
            "trailer\n<< /Root 1 0 R /Size 5 >>\nstartxref\n0\n%%EOF\n");
    }

    private static byte[] ReplaceLastStartXrefOffset(byte[] pdf, string replacement) {
        string text = Encoding.ASCII.GetString(pdf);
        int marker = text.LastIndexOf("startxref", StringComparison.Ordinal);
        Assert.True(marker >= 0);
        int start = marker + "startxref".Length;
        while (start < text.Length && char.IsWhiteSpace(text[start])) start++;
        int end = start;
        while (end < text.Length && char.IsDigit(text[end])) end++;
        return Encoding.ASCII.GetBytes(text.Substring(0, start) + replacement + text.Substring(end));
    }
}
