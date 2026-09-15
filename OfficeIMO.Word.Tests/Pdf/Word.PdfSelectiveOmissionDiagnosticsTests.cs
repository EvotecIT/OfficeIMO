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

        AssertOmission(logical, new PdfToWordOptions { ImportHeadings = false }, "PdfHeadingsNotImported");
        AssertOmission(logical, new PdfToWordOptions { ImportParagraphs = false }, "PdfParagraphsNotImported");
        AssertOmission(logical, new PdfToWordOptions { ImportLists = false }, "PdfListsNotImported");
        AssertOmission(logical, new PdfToWordOptions { ImportTables = false }, "PdfTablesNotImported");
    }

    [Fact]
    public void PdfSemanticImport_TablesOnlyReportsEveryDisabledTextFamily() {
        PdfCore.PdfDocumentReadResult logical = CreateLogicalDocumentWithAllSemanticFamilies();
        PdfWordConversionResult result = logical.ToWordDocumentResult(PdfToWordOptions.CreateTablesOnly());

        using (result.Value) {
            Assert.True(result.HasLoss);
            string[] warningCodes = result.Report.Warnings.Select(static warning => warning.Code).ToArray();
            Assert.Contains("PdfHeadingsNotImported", warningCodes);
            Assert.Contains("PdfParagraphsNotImported", warningCodes);
            Assert.Contains("PdfListsNotImported", warningCodes);
            Assert.Contains("PdfTextContentNotImported", warningCodes);
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
        PdfToWordOptions options,
        string warningCode) {
        PdfWordConversionResult result = logical.ToWordDocumentResult(options);
        using (result.Value) {
            Assert.True(result.HasLoss);
            Assert.Contains(result.Report.Warnings, warning =>
                warning.Code == warningCode &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }
    }
}
