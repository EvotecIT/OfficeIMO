using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class WordFieldUpdateReportTests {
        [Fact]
        public void Test_UpdateFieldsAndGetReport_MapsLocaleSensitiveFormulaProfile() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph()._paragraph.Append(BuildSimpleField(" = 12.5 \\# \"0.00 €\" ", "preserved"));

            WordFieldUpdateResult result = Assert.Single(document.UpdateFieldsAndGetReport().Results);

            Assert.Equal(WordFieldUpdateStatus.Unsupported, result.Status);
            Assert.Equal("FieldLocaleProfileUnsupported", result.DiagnosticCode);
            Assert.Equal(WordFieldEvaluationBasis.NotEvaluated, result.EvaluationBasis);
            Assert.Contains("locale-specific", result.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal("preserved", Assert.Single(document.InspectFields()).ResultText);
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_MapsNestedTableFormulaGeometry() {
            using WordDocument document = WordDocument.Create();
            WordTable table = document.AddTable(1, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "10";
            table.Rows[0].Cells[0]._tableCell.Append(
                new Table(
                    new TableRow(
                        new TableCell(
                            new Paragraph(new Run(new Text("99")))))));
            table.Rows[0].Cells[1].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(LEFT) ", "preserved"));

            WordFieldUpdateResult result = Assert.Single(document.UpdateFieldsAndGetReport().Results);

            Assert.Equal(WordFieldUpdateStatus.Unsupported, result.Status);
            Assert.Equal("FieldComplexTableProfileUnsupported", result.DiagnosticCode);
            Assert.Equal(WordFieldEvaluationBasis.NotEvaluated, result.EvaluationBasis);
            Assert.Contains("nested table", result.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal("preserved", Assert.Single(document.InspectFields()).ResultText);
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_MapsRelatedPartPageLayoutRequirement() {
            using WordDocument document = WordDocument.Create();
            document.AddHeadersAndFooters();
            Header header = document._wordprocessingDocument.MainDocumentPart!.HeaderParts.Single().Header!;
            header.Append(new Paragraph(BuildSimpleField(" PAGE ", "preserved")));

            WordFieldUpdateResult result = Assert.Single(document.UpdateFieldsAndGetReport().Results);

            Assert.Equal(WordFieldUpdateStatus.Skipped, result.Status);
            Assert.Equal("FieldExternalLayoutRequired", result.DiagnosticCode);
            Assert.Equal(WordFieldEvaluationBasis.ExternalLayoutRequired, result.EvaluationBasis);
            Assert.Equal(WordFieldLocationKind.Header, result.LocationKind);
            Assert.Equal("preserved", Assert.Single(document.InspectFields()).ResultText);
        }
    }
}
