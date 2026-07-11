using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfParsingModeTests {
    [Fact]
    public void LenientModeReportsAndRecoversIncorrectStreamLength() {
        byte[] pdf = BuildStreamPdf("/Length 999");

        PdfReadDocument document = PdfReadDocument.Load(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Lenient });

        PdfRepairDiagnostic repair = Assert.Single(document.RepairReport.Diagnostics, item => item.Code == "IncorrectStreamLength");
        Assert.Equal(4, repair.ObjectNumber);
        Assert.Contains("declares /Length 999", repair.Message, StringComparison.Ordinal);
        Assert.Contains("Recovered stream text", document.ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void StrictModeRejectsIncorrectStreamLengthWithStableCode() {
        byte[] pdf = BuildStreamPdf("/Length 999");

        PdfParseException exception = Assert.Throws<PdfParseException>(() =>
            PdfReadDocument.Load(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Strict }));

        Assert.Equal("IncorrectStreamLength", exception.Code);
        Assert.Equal(4, exception.ObjectNumber);
    }

    [Fact]
    public void MissingStreamLengthIsReportedOrRejectedByPolicy() {
        byte[] pdf = BuildStreamPdf(string.Empty);

        PdfReadDocument lenient = PdfReadDocument.Load(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Lenient });
        PdfParseException strict = Assert.Throws<PdfParseException>(() =>
            PdfReadDocument.Load(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Strict }));

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

        PdfReadDocument lenient = PdfReadDocument.Load(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Lenient });
        PdfParseException strict = Assert.Throws<PdfParseException>(() =>
            PdfReadDocument.Load(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Strict }));

        Assert.Contains(lenient.RepairReport.Diagnostics, item => item.Code == "MissingEndObject" && item.ObjectNumber == 3);
        Assert.Equal("MissingEndObject", strict.Code);
        Assert.Equal(3, strict.ObjectNumber);
    }

    [Fact]
    public void CleanDocumentHasEmptyRepairReport() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("Clean parse")).ToBytes();

        PdfReadDocument document = PdfReadDocument.Load(pdf, new PdfReadOptions { ParsingMode = PdfParsingMode.Strict });

        Assert.False(document.RepairReport.HasRepairs);
        Assert.Empty(document.RepairReport.Diagnostics);
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
}
