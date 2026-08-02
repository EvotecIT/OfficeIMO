using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class WordFieldUpdateReportTests {
        [Fact]
        public void Test_UpdateFieldsAndGetReport_LabelsRuntimeClockDateTimeEvidence() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph().AddField(WordFieldType.Date, customFormat: "yyyy-MM-dd");
            document.AddParagraph().AddField(WordFieldType.Time, customFormat: "HH:mm:ss");

            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

            Assert.All(report.Results, result => {
                Assert.Equal(WordFieldUpdateStatus.Updated, result.Status);
                Assert.Equal("FieldUpdatedFromRuntimeClock", result.DiagnosticCode);
                Assert.Equal(WordFieldEvaluationBasis.RuntimeClock, result.EvaluationBasis);
            });
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_LabelsCallerProvidedDateTimeEvidence() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph().AddField(WordFieldType.Date, customFormat: "yyyy-MM-dd");

            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions {
                CurrentDateTime = new DateTime(2026, 8, 2, 12, 34, 56)
            });

            WordFieldUpdateResult result = Assert.Single(report.Results);
            Assert.Equal("2026-08-02", result.ResultText);
            Assert.Equal("FieldUpdatedFromCallerDateTime", result.DiagnosticCode);
            Assert.Equal(WordFieldEvaluationBasis.CallerProvidedDateTime, result.EvaluationBasis);
        }
    }
}
