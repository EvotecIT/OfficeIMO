using OfficeIMO.Pdf.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void PdfSemanticImport_ReportsEachIndividuallyDisabledSemanticFamily() {
        PdfCore.PdfDocumentReadResult logical = CreateLogicalDocumentWithAllSemanticFamilies();
        Assert.NotEmpty(logical.Headings);
        Assert.NotEmpty(logical.Paragraphs);
        Assert.NotEmpty(logical.ListItems);
        Assert.NotEmpty(logical.Tables);

        AssertOmission(logical, new WordPdfImportOptions { ImportHeadings = false }, "pdf-headings-not-imported");
        AssertOmission(logical, new WordPdfImportOptions { ImportParagraphs = false }, "pdf-paragraphs-not-imported");
        AssertOmission(logical, new WordPdfImportOptions { ImportLists = false }, "pdf-lists-not-imported");
        AssertOmission(logical, new WordPdfImportOptions { ImportTables = false }, "pdf-tables-not-imported");
    }

    [Fact]
    public void PdfSemanticImport_TablesOnlyReportsEveryDisabledTextFamily() {
        PdfCore.PdfDocumentReadResult logical = CreateLogicalDocumentWithAllSemanticFamilies();
        PdfWordConversionResult result = logical.ConvertToWordWithEvidence(WordPdfImportOptions.CreateTablesOnly());

        using (result.Value) {
            Assert.True(result.HasLoss);
            string[] warningCodes = result.Report.Warnings.Select(static warning => warning.Code).ToArray();
            Assert.Contains("pdf-headings-not-imported", warningCodes);
            Assert.Contains("pdf-paragraphs-not-imported", warningCodes);
            Assert.Contains("pdf-lists-not-imported", warningCodes);
            Assert.Contains("pdf-text-content-not-imported", warningCodes);
        }
    }

    private static PdfCore.PdfDocumentReadResult CreateLogicalDocumentWithAllSemanticFamilies() {
        byte[] source = PdfCore.PdfDocument.Create()
            .H1("Selective heading")
            .Paragraph(paragraph => paragraph.Text("Selective paragraph"))
            .Bullets(new[] { "Selective list item" })
            .Table(new[] {
                new[] { "Item", "Amount" },
                new[] { "Service", "12.50" }
            })
            .ToBytes();
        return PdfCore.PdfDocumentReadResult.Load(source);
    }

    private static void AssertOmission(
        PdfCore.PdfDocumentReadResult logical,
        WordPdfImportOptions options,
        string warningCode) {
        PdfWordConversionResult result = logical.ConvertToWordWithEvidence(options);
        using (result.Value) {
            Assert.True(result.HasLoss);
            Assert.Contains(result.Report.Warnings, warning =>
                warning.Code == warningCode &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }
    }
}
