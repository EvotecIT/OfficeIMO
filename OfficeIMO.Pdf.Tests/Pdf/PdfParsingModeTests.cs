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

    private static byte[] BuildStreamPdf(string lengthEntry) {
        string dictionary = string.IsNullOrEmpty(lengthEntry) ? "<< >>" : "<< " + lengthEntry + " >>";
        return Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R >>\nendobj\n" +
            "4 0 obj\n" + dictionary + "\nstream\nBT (Recovered stream text) Tj ET\nendstream\nendobj\n" +
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
