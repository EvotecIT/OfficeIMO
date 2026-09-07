using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Examples.Showcase;

internal static partial class DocumentFeatureShowcase {
    private static void CreateWordExamples(string output) {
        string invoice = CreateExampleFolder(output, "word-mail-merge");
        Word.MailMerge.Example_MailMergeInvoiceWorkflow(invoice, false);
        SaveWordPreview(Path.Combine(invoice, "MailMergeInvoiceWorkflow.docx"));

        string grouped = CreateExampleFolder(output, "word-grouped-tables");
        Word.MailMerge.Example_MailMergeGroupedTableWorkflow(grouped, false);
        SaveWordPreview(Path.Combine(grouped, "MailMergeGroupedTableWorkflow.docx"));

        string review = CreateExampleFolder(output, "word-review-report");
        Word.ReviewReports.Example_ReviewReportWorkflow(review, false);
        ValidateWord(Path.Combine(review, "ReviewReportWorkflow.docx"));
        SaveMarkdownPreview(Path.Combine(review, "ReviewReportWorkflow.md"), Path.Combine(review, "review-report.pdf"));

        string comparison = CreateExampleFolder(output, "word-comparison");
        Word.CompareDocuments.Example_ReportAndRedlineWorkflow(comparison, false);
        SaveMarkdownPreview(Path.Combine(comparison, "ComparisonReportWorkflow.md"), Path.Combine(comparison, "comparison-report.pdf"));
    }

    private static void SaveWordPreview(string path) {
        using WordDocument document = WordDocument.Load(path);
        EnsureValidWord(document, path);
        document.SaveAsPdf(Path.ChangeExtension(path, ".pdf"));
    }

    private static void ValidateWord(string path) {
        using WordDocument document = WordDocument.Load(path);
        EnsureValidWord(document, path);
    }

    private static void EnsureValidWord(WordDocument document, string path) {
        var errors = document.ValidateDocument();
        if (errors.Count > 0) {
            throw new InvalidOperationException(Path.GetFileName(path) + ": " + errors[0].Description);
        }
    }
}
