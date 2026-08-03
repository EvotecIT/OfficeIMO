using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class WordFieldUpdateReportTests {
        [Fact]
        public void Test_UpdateFieldsAndGetReport_LeavesStandaloneIndexForExplicitDelegation() {
            using WordDocument document = WordDocument.Create();
            document.Settings.UpdateFieldsOnOpen = false;
            document.AddParagraph()._paragraph.Append(BuildSimpleField(" INDEX ", "stale-index"));

            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

            WordFieldUpdateResult result = Assert.Single(report.Results);
            Assert.Equal(WordFieldType.Index, result.FieldType);
            Assert.Equal(WordFieldUpdateStatus.Skipped, result.Status);
            Assert.Equal("FieldRefreshDelegated", result.DiagnosticCode);
            Assert.Equal(WordFieldEvaluationBasis.ExternalLayoutRequired, result.EvaluationBasis);
            Assert.False(document.Settings.UpdateFieldsOnOpen);
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_DoesNotActivateExternalFieldsThroughToc() {
            using WordDocument document = WordDocument.Create();
            document.Settings.UpdateFieldsOnOpen = false;
            document.AddParagraph()._paragraph.Append(
                BuildSimpleField(" TOC ", "stale-toc"));
            document.AddParagraph()._paragraph.Append(
                BuildSimpleField(
                    " INCLUDEPICTURE \\\\attacker.example\\share\\image.png ",
                    "external-placeholder"));

            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

            Assert.Contains(report.Results, result =>
                result.FieldType == WordFieldType.TOC
                && result.Status == WordFieldUpdateStatus.Skipped);
            Assert.Contains(report.Results, result =>
                result.FieldType == WordFieldType.IncludePicture
                && result.Status == WordFieldUpdateStatus.Unsupported);
            Assert.False(document.Settings.UpdateFieldsOnOpen);
        }
    }
}
