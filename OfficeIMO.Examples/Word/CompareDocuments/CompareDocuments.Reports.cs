using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class CompareDocuments {
        internal static void Example_ReportAndRedlineWorkflow(string folderPath, bool openWord) {
            Console.WriteLine("[*] Word comparison report workflow");

            string sourcePath = Path.Combine(folderPath, "ComparisonReportWorkflow.Source.docx");
            string targetPath = Path.Combine(folderPath, "ComparisonReportWorkflow.Target.docx");
            string redlinePath = Path.Combine(folderPath, "ComparisonReportWorkflow.Redline.docx");
            string inPlaceRedlinePath = Path.Combine(folderPath, "ComparisonReportWorkflow.InPlaceRedline.docx");

            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Contract comparison").Style = WordParagraphStyles.Heading1;
                document.HeaderDefaultOrCreate.AddParagraph("Classification: Internal");
                document.FooterDefaultOrCreate.AddParagraph("Contract pack: Draft");
                document.AddParagraph("Service tier: Standard");
                document.AddParagraph("The supplier must respond within 48 hours.");
                WordTable table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Incident response";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Legal";
                document.AddParagraph("Disclosure footnote").AddFootNote("Source disclosure note.");
                document.AddParagraph("Retention endnote").AddEndNote("Source retention note.");
                document.AddParagraph("Approval status: ").AddText("Manual review required");
                document.AddParagraph("Client field: ").AddField(WordFieldType.DocProperty, parameters: new List<string> { "\"Client\"" });
                document.AddStructuredDocumentTag("Northwind Traders", "Client", "ClientName");
                document.AddParagraph("Review target").AddComment("Legal Reviewer", "LR", "Confirm this wording before approval.");
                document.Save();
            }

            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Contract comparison").Style = WordParagraphStyles.Heading1;
                document.HeaderDefaultOrCreate.AddParagraph("Classification: Customer");
                document.FooterDefaultOrCreate.AddParagraph("Contract pack: Approved");
                document.AddParagraph("Service tier: Premium");
                document.AddParagraph("The supplier must respond within 24 hours.");
                WordTable table = document.AddTable(3, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Incident response";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Compliance";
                table.Rows[2].Cells[0].Paragraphs[0].Text = "Escalation review";
                table.Rows[2].Cells[1].Paragraphs[0].Text = "Support";
                WordTable auditTable = document.AddTable(1, 2);
                auditTable.Rows[0].Cells[0].Paragraphs[0].Text = "Audit pack";
                auditTable.Rows[0].Cells[1].Paragraphs[0].Text = "Quarterly";
                document.AddParagraph("Disclosure footnote").AddFootNote("Target disclosure note.");
                document.AddParagraph("Retention endnote").AddEndNote("Target retention note.");
                document.AddParagraph("Approval status: ").AddText("Manual review required").SetBold();
                document.AddParagraph("Client field: ").AddField(WordFieldType.DocProperty, parameters: new List<string> { "\"Customer\"" });
                document.AddStructuredDocumentTag("Northwind Traders", "Customer", "CustomerName");
                document.AddParagraph("Review target").AddComment("Legal Reviewer", "LR", "Approved for the premium contract.");
                document.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);
            EnsureComparisonFinding(result, WordComparisonScope.Run, "Run formatting changed.");
            EnsureComparisonFinding(result, WordComparisonScope.Field, "Field changed.");
            EnsureComparisonFinding(result, WordComparisonScope.ContentControl, "Content control changed.");
            EnsureComparisonFinding(result, WordComparisonScope.TableCell, "Table cell text changed.");
            EnsureComparisonFinding(result, WordComparisonScope.TableRow, "Table row inserted.");
            EnsureComparisonFinding(result, WordComparisonScope.Table, "Table inserted.");
            EnsureComparisonFinding(result, WordComparisonScope.Paragraph, "Paragraph text changed.");
            File.WriteAllText(Path.Combine(folderPath, "ComparisonReportWorkflow.json"), result.ToJson(), Encoding.UTF8);
            File.WriteAllText(Path.Combine(folderPath, "ComparisonReportWorkflow.md"), result.ToMarkdown(), Encoding.UTF8);
            File.WriteAllText(Path.Combine(folderPath, "ComparisonReportWorkflow.txt"), result.ToTextSummary(), Encoding.UTF8);

            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                redlinePath,
                new WordComparisonRedlineOptions {
                    Author = "OfficeIMO Examples",
                    DateTime = new DateTime(2026, 6, 29, 12, 0, 0, DateTimeKind.Utc),
                    TrackReviewFindings = false,
                    TrackFormattingFindings = false
                });

            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                inPlaceRedlinePath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Examples",
                    DateTime = new DateTime(2026, 6, 29, 12, 0, 0, DateTimeKind.Utc),
                    TrackReviewFindings = false,
                    TrackFormattingFindings = false
                });

            using (WordDocument redline = WordDocument.Load(redlinePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                var errors = redline.ValidateDocument();
                if (errors.Count > 0) {
                    throw new InvalidOperationException("Generated comparison redline document failed Open XML validation: " + errors[0].Description);
                }
            }

            using (WordDocument inPlaceRedline = WordDocument.Load(inPlaceRedlinePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                var errors = inPlaceRedline.ValidateDocument();
                if (errors.Count > 0) {
                    throw new InvalidOperationException("Generated in-place comparison redline document failed Open XML validation: " + errors[0].Description);
                }
            }

            if (openWord) {
                using WordDocument inPlaceRedline = WordDocument.Load(inPlaceRedlinePath);
                inPlaceRedline.Save(new WordSaveOptions { OpenAfterSave = true });
            }
        }

        private static void EnsureComparisonFinding(WordComparisonResult result, WordComparisonScope scope, string message) {
            foreach (WordComparisonFinding finding in result.Findings) {
                if (finding.Scope == scope && string.Equals(finding.Message, message, StringComparison.Ordinal)) {
                    return;
                }
            }

            throw new InvalidOperationException($"Comparison report example did not produce the expected {scope} finding: {message}");
        }
    }
}
