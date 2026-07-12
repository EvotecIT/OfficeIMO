using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class MarketReadinessProofGallery {
        private static void CreatePremiumWorkflowReportsProof(string galleryPath) {
            string scenarioPath = Path.Combine(galleryPath, "05-premium-workflow-reports");
            Directory.CreateDirectory(scenarioPath);

            CreateFeaturePreflightProof(scenarioPath);
            CreateReviewReportProof(scenarioPath);
            CreateComparisonReportProof(scenarioPath);
            CreateFieldRefreshReportProof(scenarioPath);
            CreateTemplatePreflightProof(scenarioPath);
            CreateSignaturePreflightProof(scenarioPath);
            WriteValidationReport(scenarioPath);
        }

        private static void CreateFeaturePreflightProof(string scenarioPath) {
            string featurePath = Path.Combine(scenarioPath, "unknown-document-preflight.docx");
            using (WordDocument document = WordDocument.Create(featurePath)) {
                document.BuiltinDocumentProperties.Title = "Unknown Contract Package";
                document.CustomDocumentProperties["Client"] = new WordCustomProperty("Northwind Traders");

                document.AddParagraph("Unknown Contract Package").Style = WordParagraphStyles.Heading1;
                document.AddParagraph("External policy reference: ")
                    .AddHyperLink("incident policy", new Uri("https://example.com/policies/incident-response"));
                document.AddParagraph("Approver").AddCheckBox(true, "Approval", "ApprovalTag");
                document.AddParagraph("Client: ")
                    .AddField(WordFieldType.DocProperty, parameters: new List<string> { "\"Client\"" });

                WordTable table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Incident notification";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Legal";

                WordParagraph reviewTarget = document.AddParagraph("Notification language needs legal review.");
                reviewTarget.AddComment("Legal Reviewer", "LR", "Confirm this against the customer policy.");
                document.Save();
            }

            AddFeaturePreflightPackageSignals(featurePath);
            PremiumWorkflowExampleUtilities.AddSyntheticSignatureMetadata(featurePath);

            using (WordDocument document = WordDocument.Load(featurePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordFeatureReport report = document.InspectFeatures();
                File.WriteAllText(Path.Combine(scenarioPath, "feature-report.md"), report.ToMarkdown(), Encoding.UTF8);
                File.WriteAllText(Path.Combine(scenarioPath, "feature-report.json"), SerializeFeatureReport(report), Encoding.UTF8);
            }
        }

        private static void CreateReviewReportProof(string scenarioPath) {
            string reviewPath = Path.Combine(scenarioPath, "reviewed-contract.docx");
            using (WordDocument document = WordDocument.Create(reviewPath)) {
                document.AddParagraph("Service Agreement").Style = WordParagraphStyles.Heading1;
                WordParagraph scope = document.AddParagraph("The supplier must notify the customer within 48 hours.");
                scope.AddComment("Legal Reviewer", "LR", "Please align this notification period with the incident policy.");
                WordComment comment = WordComment.GetAllComments(document).Last();
                comment.AddReply("Document Owner", "DO", "Updated in the next revision.");
                comment.MarkResolved();

                WordParagraph tracked = document.AddParagraph("Tracked language: ");
                tracked.AddDeletedText("best effort", "Legal Reviewer", new DateTime(2026, 6, 1, 9, 0, 0, DateTimeKind.Utc));
                tracked.AddInsertedText("commercially reasonable efforts", "Legal Reviewer", new DateTime(2026, 6, 1, 9, 5, 0, DateTimeKind.Utc));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(reviewPath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewReport report = document.InspectReviewReport();
                File.WriteAllText(Path.Combine(scenarioPath, "review-report.md"), report.ToMarkdown(), Encoding.UTF8);
                File.WriteAllText(Path.Combine(scenarioPath, "review-report.json"), report.ToJson(), Encoding.UTF8);
            }
        }

        private static void CreateComparisonReportProof(string scenarioPath) {
            string sourcePath = Path.Combine(scenarioPath, "comparison-source.docx");
            using (WordDocument source = WordDocument.Create(sourcePath)) {
                source.BuiltinDocumentProperties.Title = "Agreement Draft";
                source.AddParagraph("Agreement Draft").Style = WordParagraphStyles.Heading1;
                source.HeaderDefaultOrCreate.AddParagraph("Classification: Internal");
                source.FooterDefaultOrCreate.AddParagraph("Agreement pack: Draft");
                source.AddParagraph("Payment is due within 30 days.");
                source.AddParagraph("Service credits are capped at one month of fees.");
                WordTable table = source.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Payment terms";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Finance";
                source.AddParagraph("Disclosure footnote").AddFootNote("Source disclosure note.");
                source.AddParagraph("Retention endnote").AddEndNote("Source retention note.");
                source.AddParagraph("Client: ").AddField(WordFieldType.DocProperty, parameters: new List<string> { "\"Client\"" });
                source.CustomDocumentProperties["Client"] = new WordCustomProperty("Northwind Traders");
                source.UpdateFields();
                source.Save();
            }

            string targetPath = Path.Combine(scenarioPath, "comparison-target.docx");
            using (WordDocument target = WordDocument.Create(targetPath)) {
                target.BuiltinDocumentProperties.Title = "Agreement Draft";
                target.AddParagraph("Agreement Draft").Style = WordParagraphStyles.Heading1;
                target.HeaderDefaultOrCreate.AddParagraph("Classification: Customer");
                target.FooterDefaultOrCreate.AddParagraph("Agreement pack: Approved");
                target.AddParagraph("Payment is due within 14 days.");
                target.AddParagraph("Service credits are capped at two months of fees.");
                target.AddParagraph("Escalation contacts must be reviewed quarterly.");
                WordTable table = target.AddTable(3, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Payment terms";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Legal";
                table.Rows[2].Cells[0].Paragraphs[0].Text = "Escalation review";
                table.Rows[2].Cells[1].Paragraphs[0].Text = "Support";
                WordTable auditTable = target.AddTable(1, 2);
                auditTable.Rows[0].Cells[0].Paragraphs[0].Text = "Audit pack";
                auditTable.Rows[0].Cells[1].Paragraphs[0].Text = "Quarterly";
                target.AddParagraph("Disclosure footnote").AddFootNote("Target disclosure note.");
                target.AddParagraph("Retention endnote").AddEndNote("Target retention note.");
                target.AddParagraph("Client: ").AddField(WordFieldType.DocProperty, parameters: new List<string> { "\"Client\"" });
                target.CustomDocumentProperties["Client"] = new WordCustomProperty("Contoso Legal");
                target.UpdateFields();
                target.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(
                sourcePath,
                targetPath,
                new WordComparisonOptions {
                    CompareFields = true,
                    CompareRunFormatting = true,
                    IgnoreWhitespace = true
                });

            File.WriteAllText(Path.Combine(scenarioPath, "comparison-report.md"), WordComparisonReportWriter.ToMarkdown(result), Encoding.UTF8);
            File.WriteAllText(Path.Combine(scenarioPath, "comparison-report.json"), WordComparisonReportWriter.ToJson(result), Encoding.UTF8);
            File.WriteAllText(Path.Combine(scenarioPath, "comparison-summary.txt"), WordComparisonReportWriter.ToTextSummary(result), Encoding.UTF8);

            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                Path.Combine(scenarioPath, "comparison-redline.docx"),
                new WordComparisonRedlineOptions {
                    Author = "OfficeIMO Example",
                    DateTime = new DateTime(2026, 6, 1, 10, 0, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        CompareFields = true,
                        CompareRunFormatting = true,
                        IgnoreWhitespace = true
                    }
                });

            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                Path.Combine(scenarioPath, "comparison-redline-in-place.docx"),
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Example",
                    DateTime = new DateTime(2026, 6, 1, 10, 0, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        CompareFields = true,
                        CompareRunFormatting = true,
                        IgnoreWhitespace = true
                    },
                    TrackFormattingFindings = false
                });
        }

        private static void CreateFieldRefreshReportProof(string scenarioPath) {
            string fieldPath = Path.Combine(scenarioPath, "field-refresh-report.docx");
            using (WordDocument document = WordDocument.Create(fieldPath)) {
                document.BuiltinDocumentProperties.Creator = "OfficeIMO";
                document.BuiltinDocumentProperties.Title = "Quarterly Controls Report";
                document.CustomDocumentProperties["Client"] = new WordCustomProperty("Northwind Traders");

                document.AddParagraph("Quarterly Controls Report").Style = WordParagraphStyles.Heading1;
                document.AddParagraph("Executive Summary").AddBookmark("ExecutiveSummary");
                document.AddParagraph("Prepared by: ").AddField(WordFieldType.Author);
                document.AddParagraph("Client: ").AddField(WordFieldType.DocProperty, parameters: new List<string> { "\"Client\"" });
                document.AddPageBreak();
                document.AddParagraph("Summary reference: ").AddField(WordFieldType.Ref, parameters: new List<string> { "ExecutiveSummary" });
                document.AddParagraph("Summary page: ").AddField(WordFieldType.PageRef, parameters: new List<string> { "ExecutiveSummary" });

                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();
                WriteFieldUpdateReport(Path.Combine(scenarioPath, "field-refresh-report.md"), report);
                document.Save();
            }
        }

        private static void CreateSignaturePreflightProof(string scenarioPath) {
            string signaturePath = Path.Combine(scenarioPath, "signature-preflight.docx");
            using (WordDocument document = WordDocument.Create(signaturePath)) {
                document.AddParagraph("Signed package metadata preflight").Style = WordParagraphStyles.Heading1;
                document.AddParagraph("OfficeIMO can inspect signature package metadata and block accidental saves by default.");
                document.Save();
            }

            PremiumWorkflowExampleUtilities.AddSyntheticSignatureMetadata(signaturePath);

            WordSignatureValidationReport validationReport;
            using (WordDocument document = WordDocument.Load(signaturePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                validationReport = document.ValidateSignatures();
            }

            string savePolicyMessage;
            using (WordDocument document = WordDocument.Load(signaturePath)) {
                document.AddParagraph("This edit is intentionally blocked by the default signed-document save policy.");
                try {
                    document.Save();
                    savePolicyMessage = "Save unexpectedly succeeded.";
                } catch (WordSignatureSavePolicyException ex) {
                    savePolicyMessage = ex.Message;
                }
            }

            PremiumWorkflowExampleUtilities.WriteSignaturePreflightReport(Path.Combine(scenarioPath, "signature-preflight.md"), validationReport, savePolicyMessage);
        }

        private static void CreateTemplatePreflightProof(string scenarioPath) {
            string templatePath = Path.Combine(scenarioPath, "template-preflight.docx");
            using (WordDocument document = WordDocument.Create(templatePath)) {
                document.AddParagraph("Proposal Template").Style = WordParagraphStyles.Heading1;
                document.AddParagraph("Client: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"ClientName\"" });
                document.AddParagraph("{{#ShowDiscount}}");
                document.AddParagraph("Discount: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Discount\"" });
                document.AddParagraph("{{/ShowDiscount}}");
                document.AddParagraph("{{#each Services}}");
                document.AddParagraph("Service: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"ServiceName\"" });
                document.AddParagraph("{{/each Services}}");

                WordTemplatePreflightReport cleanReport = WordMailMerge.PreflightTemplate(
                    document,
                    mergeFieldNames: new[] { "ClientName", "Discount", "ServiceName" },
                    conditionNames: new[] { "ShowDiscount" },
                    repeatingBlockNames: new[] { "Services" });
                File.WriteAllText(Path.Combine(scenarioPath, "template-preflight.md"), cleanReport.ToMarkdown(), Encoding.UTF8);
                File.WriteAllText(Path.Combine(scenarioPath, "template-preflight.json"), cleanReport.ToJson(), Encoding.UTF8);

                WordTemplatePreflightReport blockedReport = WordMailMerge.PreflightTemplate(
                    document,
                    mergeFieldNames: new[] { "ClientName" },
                    conditionNames: Array.Empty<string>(),
                    repeatingBlockNames: Array.Empty<string>());
                File.WriteAllText(Path.Combine(scenarioPath, "template-preflight-blocked.md"), blockedReport.ToMarkdown(), Encoding.UTF8);
                File.WriteAllText(Path.Combine(scenarioPath, "template-preflight-blocked.json"), blockedReport.ToJson(), Encoding.UTF8);

                document.Save();
            }
        }

        private static void WriteFieldUpdateReport(string path, WordFieldUpdateReport report) {
            var builder = new StringBuilder();
            builder.AppendLine("# Field Refresh Report");
            builder.AppendLine();
            builder.AppendLine("- Total fields: " + report.TotalCount);
            builder.AppendLine("- Updated: " + report.UpdatedCount);
            builder.AppendLine("- Skipped: " + report.SkippedCount);
            builder.AppendLine("- Unsupported: " + report.UnsupportedCount);
            builder.AppendLine("- Parse errors: " + report.ParseErrorCount);
            builder.AppendLine();
            builder.AppendLine("| # | Type | Status | Result | Message |");
            builder.AppendLine("| ---: | --- | --- | --- | --- |");
            foreach (WordFieldUpdateResult result in report.Results) {
                builder.Append("| ");
                builder.Append(result.Index);
                builder.Append(" | ");
                builder.Append(result.FieldType?.ToString() ?? "Unknown");
                builder.Append(" | ");
                builder.Append(result.Status);
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(result.ResultText));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(result.Message));
                builder.AppendLine(" |");
            }

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        private static string EscapeMarkdownCell(string? value) {
            return (value ?? string.Empty)
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("|", "\\|");
        }

        private static string SerializeFeatureReport(WordFeatureReport report) {
            var payload = new {
                totalFindings = report.Features.Count,
                editableFeatures = report.EditableFeatures.Count,
                partiallyEditableFeatures = report.PartiallyEditableFeatures.Count,
                preservedFeatures = report.PreservedFeatures.Count,
                unsupportedFeatures = report.UnsupportedFeatures.Count,
                hasAdvancedFeatures = report.HasAdvancedFeatures,
                features = report.Features.Select(feature => new {
                    feature.Category,
                    feature.Name,
                    supportLevel = feature.SupportLevel.ToString(),
                    feature.Count,
                    feature.Scope,
                    feature.Note,
                    feature.Details
                })
            };

            return JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true });
        }

        private static void AddFeaturePreflightPackageSignals(string filePath) {
            using WordprocessingDocument package = WordprocessingDocument.Open(filePath, true);
            MainDocumentPart mainPart = package.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart is missing.");
            AddExtendedPart(mainPart,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control",
                "application/vnd.ms-office.activeX+xml",
                "<ax:ocx xmlns:ax=\"http://schemas.microsoft.com/office/2006/activeX\" />");
        }

        private static void AddExtendedPart(OpenXmlPartContainer container, string relationshipType, string contentType, string xml) {
            ExtendedPart part = container.AddExtendedPart(relationshipType, contentType, "xml");
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(xml));
            part.FeedData(stream);
        }

    }
}
