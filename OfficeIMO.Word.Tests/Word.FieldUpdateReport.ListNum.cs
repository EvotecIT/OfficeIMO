using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class WordFieldUpdateReportTests {
        [Theory]
        [InlineData("NumberDefault", "1)|a)|i)|(1)|(a)|(i)|1.|a.|i.")]
        [InlineData("OutlineDefault", "I.|A.|1.|a)|(1)|(a)|(i)|(a)|(i)")]
        [InlineData("LegalDefault", "1.|1.1.|1.1.1.|1.1.1.1.|1.1.1.1.1.|1.1.1.1.1.1.|1.1.1.1.1.1.1.|1.1.1.1.1.1.1.1.|1.1.1.1.1.1.1.1.1.")]
        public void Test_UpdateFieldsAndGetReport_EvaluatesBuiltInListNumProfiles(string profile, string expectedText) {
            string filePath = Path.Combine(_directoryWithFiles, $"FieldUpdate.ListNum.{profile}.docx");
            string[] expected = expectedText.Split('|');

            using (WordDocument document = WordDocument.Create(filePath)) {
                for (int level = 1; level <= 9; level++) {
                    document.AddParagraph()._paragraph.Append(BuildSimpleField($" LISTNUM {profile} \\l {level} ", "stale"));
                }

                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(9, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(expected, report.Results.Select(result => result.ResultText).ToArray());
                Assert.All(report.Results, result => {
                    Assert.Equal("FieldUpdatedInvariant", result.DiagnosticCode);
                    Assert.Equal(WordFieldEvaluationBasis.InvariantDocumentModel, result.EvaluationBasis);
                });
                document.Save();
            }

            using WordDocument loaded = WordDocument.Load(filePath);
            Assert.Equal(expected, loaded.InspectFields().Select(field => field.ResultText).ToArray());
            Assert.Empty(new OpenXmlValidator().Validate(loaded._wordprocessingDocument));
        }

        [Theory]
        [InlineData("NumberDefault", "1)|a)|b)|2)|a)|i)|3)|i)")]
        [InlineData("OutlineDefault", "I.|A.|B.|II.|A.|1.|III.|1.")]
        [InlineData("LegalDefault", "1.|1.1.|1.2.|2.|2.1.|2.1.1.|3.|3.1.1.")]
        public void Test_UpdateFieldsAndGetReport_TracksListNumLevelsAndResetsDescendants(string profile, string expectedText) {
            using WordDocument document = WordDocument.Create();
            int[] levels = { 1, 2, 2, 1, 2, 3, 1, 3 };
            foreach (int level in levels) {
                document.AddParagraph()._paragraph.Append(BuildSimpleField($" LISTNUM {profile} \\l {level} ", "stale"));
            }

            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

            Assert.Equal(expectedText.Split('|'), report.Results.Select(result => result.ResultText).ToArray());
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_AppliesListNumStartSwitchAndParagraphLevel() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph()._paragraph.Append(BuildSimpleField(" LISTNUM LegalDefault \\l 1 \\s 5 ", "stale"));
            document.AddParagraph()._paragraph.Append(BuildSimpleField(" LISTNUM LegalDefault \\l 2 ", "stale"));

            WordParagraph paragraphLevel = document.AddParagraph();
            paragraphLevel._paragraph.ParagraphProperties = new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = 4 },
                    new NumberingId { Val = 1 }));
            paragraphLevel._paragraph.Append(BuildSimpleField(" LISTNUM NumberDefault ", "stale"));

            document.AddParagraph()._paragraph.Append(BuildSimpleField(" LISTNUM ", "stale"));

            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

            Assert.Equal(new[] { "5.", "5.1.", "(a)", "1)" }, report.Results.Select(result => result.ResultText).ToArray());
        }

        [Theory]
        [InlineData(" LISTNUM CustomTemplate ", "FieldListNumberingProfileUnsupported")]
        [InlineData(" LISTNUM NumberDefault \\l 10 ", "FieldListNumberingProfileUnsupported")]
        [InlineData(" LISTNUM NumberDefault \\s 0 ", "FieldListNumberingProfileUnsupported")]
        [InlineData(" LISTNUM NumberDefault \\x 1 ", "FieldListNumberingProfileUnsupported")]
        public void Test_UpdateFieldsAndGetReport_PreservesUnsupportedListNumProfiles(string instruction, string diagnosticCode) {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph()._paragraph.Append(BuildSimpleField(instruction, "preserved"));

            WordFieldUpdateResult result = Assert.Single(document.UpdateFieldsAndGetReport().Results);

            Assert.Equal(WordFieldUpdateStatus.Unsupported, result.Status);
            Assert.Equal(diagnosticCode, result.DiagnosticCode);
            Assert.Equal(WordFieldEvaluationBasis.NotEvaluated, result.EvaluationBasis);
            Assert.Equal("preserved", Assert.Single(document.InspectFields()).ResultText);
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_DiagnosesNestedListNumWithoutUpdatingIt() {
            using WordDocument document = WordDocument.Create();
            var nested = BuildSimpleField(" LISTNUM NumberDefault ", "nested-preserved");
            var outer = new SimpleField(
                new Run(new Text("outer prefix ") { Space = SpaceProcessingModeValues.Preserve }),
                nested,
                new Run(new Text(" outer suffix") { Space = SpaceProcessingModeValues.Preserve })) {
                Instruction = " QUOTE \"outer\" "
            };
            document.AddParagraph()._paragraph.Append(outer);

            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();
            WordFieldUpdateResult nestedResult = Assert.Single(report.Results, result => result.FieldType == WordFieldType.ListNum);

            Assert.Equal(WordFieldUpdateStatus.Unsupported, nestedResult.Status);
            Assert.Equal("FieldNestedInstructionUnsupported", nestedResult.DiagnosticCode);
            Assert.Contains("ignored by Word", nestedResult.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_DiagnosesListNumCounterOverflowWithoutThrowing() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph()._paragraph.Append(BuildSimpleField($" LISTNUM NumberDefault \\s {int.MaxValue} ", "stale-start"));
            document.AddParagraph()._paragraph.Append(BuildSimpleField(" LISTNUM NumberDefault ", "preserved"));

            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

            Assert.Equal(int.MaxValue.ToString(System.Globalization.CultureInfo.InvariantCulture) + ")", report.Results[0].ResultText);
            Assert.Equal(WordFieldUpdateStatus.Unsupported, report.Results[1].Status);
            Assert.Equal("FieldListNumberingProfileUnsupported", report.Results[1].DiagnosticCode);
            Assert.Equal("preserved", document.InspectFields()[1].ResultText);
        }
    }
}
