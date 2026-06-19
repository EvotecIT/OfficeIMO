using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Examples.Word {
    internal static class MarketReadinessProofGallery {
        public static void Example_GenerateWordMarketReadinessProof(string folderPath, bool openWord) {
            Console.WriteLine("[*] Word market readiness proof gallery");

            string galleryPath = Path.Combine(folderPath, "WordMarketReadinessProof");
            Directory.CreateDirectory(galleryPath);

            CreateTemplateAssemblyProof(galleryPath);
            CreateReviewDiffProof(galleryPath);
            CreateHtmlProof(galleryPath);
            CreateMarkdownProof(galleryPath);

            Console.WriteLine($"Proof gallery written to: {galleryPath}");
        }

        private static void CreateTemplateAssemblyProof(string galleryPath) {
            string scenarioPath = Path.Combine(galleryPath, "01-template-assembly");
            Directory.CreateDirectory(scenarioPath);

            string templatePath = Path.Combine(scenarioPath, "status-report-template.docx");
            using (WordDocument template = WordDocument.Create(templatePath)) {
                template.AddParagraph("Project Status Report").Style = WordParagraphStyles.Heading1;
                template.AddParagraph("Project: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"ProjectName\"" });
                template.AddParagraph("Owner: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Owner\"" });
                template.AddParagraph("Status: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Status\"" });
                template.AddParagraph("Next step: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"NextStep\"" });

                WordMailMergeTemplateInspection inspection = WordMailMerge.InspectTemplate(
                    template,
                    new[] { "ProjectName", "Owner", "Status", "NextStep" });

                WriteTemplateDiagnostics(Path.Combine(scenarioPath, "template-diagnostics.md"), inspection);
                template.Save();
            }

            var records = new[] {
                new Dictionary<string, string> {
                    { "ProjectName", "OfficeIMO.Word readiness" },
                    { "Owner", "Document automation team" },
                    { "Status", "Green" },
                    { "NextStep", "Publish proof gallery outputs" }
                },
                new Dictionary<string, string> {
                    { "ProjectName", "Contract review workflow" },
                    { "Owner", "Legal operations" },
                    { "Status", "Amber" },
                    { "NextStep", "Add run-level diff reporting" }
                }
            };

            WriteRecords(Path.Combine(scenarioPath, "data-input.md"), records);
            WordMailMerge.ExecuteBatch(
                templatePath,
                records,
                (index, _) => Path.Combine(scenarioPath, $"status-report-{index + 1}.docx"));

            WriteValidationReport(scenarioPath);
        }

        private static void CreateReviewDiffProof(string galleryPath) {
            string scenarioPath = Path.Combine(galleryPath, "02-review-diff");
            Directory.CreateDirectory(scenarioPath);

            string sourcePath = Path.Combine(scenarioPath, "policy-source.docx");
            using (WordDocument source = WordDocument.Create(sourcePath)) {
                source.AddParagraph("Acceptable Use Policy").Style = WordParagraphStyles.Heading1;
                source.AddParagraph("Remote access requires quarterly approval.");
                source.AddParagraph("Audit logs must be retained for 90 days.");
                source.AddParagraph("Both parties accept the control language.");
                WordTable table = source.AddTable(3, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "MFA";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Security";
                table.Rows[2].Cells[0].Paragraphs[0].Text = "Logging";
                table.Rows[2].Cells[1].Paragraphs[0].Text = "Platform";
                source.AddParagraph().AddImage(Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png"));
                source.AddParagraph().AddImage(Path.Combine(AppContext.BaseDirectory, "Images", "snail.bmp"));
                source.Save();
            }

            string targetPath = Path.Combine(scenarioPath, "policy-target.docx");
            using (WordDocument target = WordDocument.Create(targetPath)) {
                target.AddParagraph("Acceptable Use Policy").Style = WordParagraphStyles.Heading1;
                target.AddParagraph("Remote access requires monthly approval.");
                target.AddParagraph("Privileged access requires manager attestation.");
                target.AddParagraph("Audit logs must be retained for 90 days.");
                target.AddParagraph("Both parties accept the control language.");
                WordTable table = target.AddTable(4, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "MFA";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Identity";
                table.Rows[2].Cells[0].Paragraphs[0].Text = "Review";
                table.Rows[2].Cells[1].Paragraphs[0].Text = "Compliance";
                table.Rows[3].Cells[0].Paragraphs[0].Text = "Logging";
                table.Rows[3].Cells[1].Paragraphs[0].Text = "Platform";
                target.AddParagraph().AddImage(Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png"));
                target.AddParagraph().AddImage(Path.Combine(AppContext.BaseDirectory, "Images", "BackgroundImage.png"));
                target.AddParagraph().AddImage(Path.Combine(AppContext.BaseDirectory, "Images", "snail.bmp"));
                target.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);
            WriteComparisonReport(Path.Combine(scenarioPath, "structured-diff.md"), result);
            WriteValidationReport(scenarioPath);
        }

        private static void CreateHtmlProof(string galleryPath) {
            string scenarioPath = Path.Combine(galleryPath, "03-html-conversion");
            Directory.CreateDirectory(scenarioPath);

            var scenarios = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
                ["browser-copy"] = "<h1>Incident Summary</h1><p><strong>Severity:</strong> High</p><ul><li>Detected by monitoring</li><li>Owner assigned</li></ul>",
                ["cms-article"] = "<article><h1>Release Notes</h1><p>The release improves document review workflows.</p><table><tr><th>Area</th><th>Status</th></tr><tr><td>Word</td><td>Ready</td></tr></table></article>"
            };

            foreach (KeyValuePair<string, string> scenario in scenarios) {
                string sourcePath = Path.Combine(scenarioPath, scenario.Key + ".html");
                string docxPath = Path.Combine(scenarioPath, scenario.Key + ".docx");
                string roundTripPath = Path.Combine(scenarioPath, scenario.Key + ".roundtrip.html");
                string diagnosticsPath = Path.Combine(scenarioPath, scenario.Key + ".diagnostics.md");

                File.WriteAllText(sourcePath, scenario.Value, Encoding.UTF8);

                HtmlToWordOptions options = HtmlToWordOptions.CreateUntrustedHtmlProfile();
                using WordDocument document = scenario.Value.LoadFromHtml(options);
                document.Save(docxPath);
                File.WriteAllText(roundTripPath, document.ToHtml(new WordToHtmlOptions { IncludeDefaultCss = true }), Encoding.UTF8);
                WriteHtmlDiagnostics(diagnosticsPath, options.Diagnostics);
            }

            WriteValidationReport(scenarioPath);
        }

        private static void CreateMarkdownProof(string galleryPath) {
            string scenarioPath = Path.Combine(galleryPath, "04-markdown-conversion");
            Directory.CreateDirectory(scenarioPath);

            var scenarios = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
                ["status-report"] = "# Status Report\n\nOwner: Document automation team\n\n- Green: template inspection\n- Amber: run-level review diff\n\n| Area | State |\n| --- | --- |\n| Word | Ready |\n",
                ["developer-doc"] = "# Developer Note\n\nUse `WordDocumentComparer.CompareStructure` to produce structured findings.\n\n1. Create source document\n2. Create target document\n3. Review findings\n"
            };

            foreach (KeyValuePair<string, string> scenario in scenarios) {
                string sourcePath = Path.Combine(scenarioPath, scenario.Key + ".md");
                string docxPath = Path.Combine(scenarioPath, scenario.Key + ".docx");
                string roundTripPath = Path.Combine(scenarioPath, scenario.Key + ".roundtrip.md");
                string diagnosticsPath = Path.Combine(scenarioPath, scenario.Key + ".diagnostics.md");
                var warnings = new List<string>();

                File.WriteAllText(sourcePath, scenario.Value, Encoding.UTF8);

                var toWordOptions = new MarkdownToWordOptions {
                    OnWarning = warnings.Add,
                    PreferNarrativeSingleLineDefinitions = true
                };
                using WordDocument document = scenario.Value.LoadFromMarkdown(toWordOptions);
                document.Save(docxPath);

                var toMarkdownOptions = new WordToMarkdownOptions {
                    OnWarning = warnings.Add
                };
                File.WriteAllText(roundTripPath, document.ToMarkdown(toMarkdownOptions), Encoding.UTF8);
                WriteWarnings(diagnosticsPath, warnings);
            }

            WriteValidationReport(scenarioPath);
        }

        private static void WriteTemplateDiagnostics(string path, WordMailMergeTemplateInspection inspection) {
            var builder = new StringBuilder();
            builder.AppendLine("# Template Diagnostics");
            builder.AppendLine();
            builder.AppendLine("- Valid: " + inspection.IsValid);
            builder.AppendLine("- Merge fields: " + string.Join(", ", inspection.MergeFieldNames));
            builder.AppendLine("- Issues: " + inspection.Issues.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
            foreach (WordMailMergeTemplateIssue issue in inspection.Issues) {
                builder.AppendLine("  - " + issue.Kind + ": " + issue.Message);
            }

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        private static void WriteRecords(string path, IEnumerable<IDictionary<string, string>> records) {
            var builder = new StringBuilder();
            builder.AppendLine("# Data Input");
            int index = 1;
            foreach (IDictionary<string, string> record in records) {
                builder.AppendLine();
                builder.AppendLine("## Record " + index.ToString(System.Globalization.CultureInfo.InvariantCulture));
                foreach (KeyValuePair<string, string> value in record) {
                    builder.AppendLine("- " + value.Key + ": " + value.Value);
                }

                index++;
            }

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        private static void WriteComparisonReport(string path, WordComparisonResult result) {
            var builder = new StringBuilder();
            builder.AppendLine("# Structured Diff");
            builder.AppendLine();
            builder.AppendLine("- Source: " + result.SourcePath);
            builder.AppendLine("- Target: " + result.TargetPath);
            builder.AppendLine("- Has changes: " + result.HasChanges);
            builder.AppendLine();
            foreach (WordComparisonFinding finding in result.Findings) {
                builder.AppendLine("- " + finding.ChangeKind + " " + finding.Scope + " at `" + finding.Location + "`"
                    + " (source index: " + FormatIndex(finding.SourceIndex) + ", target index: " + FormatIndex(finding.TargetIndex) + "): "
                    + finding.Message);
            }

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        private static string FormatIndex(int? index) {
            return index.HasValue ? index.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : "-";
        }

        private static void WriteHtmlDiagnostics(string path, IReadOnlyList<HtmlConversionDiagnostic> diagnostics) {
            var builder = new StringBuilder();
            builder.AppendLine("# HTML Conversion Diagnostics");
            builder.AppendLine();
            if (diagnostics.Count == 0) {
                builder.AppendLine("- No diagnostics.");
            } else {
                foreach (HtmlConversionDiagnostic diagnostic in diagnostics) {
                    builder.AppendLine("- " + diagnostic.Severity + " " + diagnostic.Code + ": " + diagnostic.Message);
                }
            }

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        private static void WriteWarnings(string path, IReadOnlyList<string> warnings) {
            var builder = new StringBuilder();
            builder.AppendLine("# Conversion Diagnostics");
            builder.AppendLine();
            if (warnings.Count == 0) {
                builder.AppendLine("- No diagnostics.");
            } else {
                foreach (string warning in warnings) {
                    builder.AppendLine("- " + warning);
                }
            }

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        private static void WriteValidationReport(string scenarioPath) {
            var builder = new StringBuilder();
            builder.AppendLine("# Open XML Validation");
            builder.AppendLine();

            foreach (string docxPath in Directory.GetFiles(scenarioPath, "*.docx").OrderBy(path => path, StringComparer.OrdinalIgnoreCase)) {
                IReadOnlyList<ValidationErrorInfo> errors = ValidateDocx(docxPath);
                builder.AppendLine("- " + Path.GetFileName(docxPath) + ": " + errors.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + " validation errors");
                foreach (ValidationErrorInfo error in errors) {
                    builder.AppendLine("  - " + error.Description);
                }
            }

            File.WriteAllText(Path.Combine(scenarioPath, "openxml-validation.md"), builder.ToString(), Encoding.UTF8);
        }

        private static IReadOnlyList<ValidationErrorInfo> ValidateDocx(string path) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, false);
            var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
            return validator.Validate(document).ToList();
        }
    }
}
