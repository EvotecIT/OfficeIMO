using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class WordFieldUpdateReportTests {
        [Fact]
        public void Test_UpdateFieldsAndGetReport_QueuesStandaloneIndexForUpdateOnOpen() {
            using WordDocument document = WordDocument.Create();
            document.Settings.UpdateFieldsOnOpen = false;
            document.AddParagraph()._paragraph.Append(BuildSimpleField(" INDEX ", "stale-index"));

            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

            WordFieldUpdateResult result = Assert.Single(report.Results);
            Assert.Equal(WordFieldType.Index, result.FieldType);
            Assert.Equal(WordFieldUpdateStatus.Skipped, result.Status);
            Assert.Equal("FieldRefreshDelegated", result.DiagnosticCode);
            Assert.Equal(WordFieldEvaluationBasis.ExternalLayoutRequired, result.EvaluationBasis);
            Assert.True(document.Settings.UpdateFieldsOnOpen);
        }
    }
}
