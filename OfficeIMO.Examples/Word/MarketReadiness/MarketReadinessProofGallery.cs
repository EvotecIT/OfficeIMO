using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Examples.Word {
    internal static partial class MarketReadinessProofGallery {
        public static void Example_GenerateWordMarketReadinessProof(string folderPath, bool openWord) {
            Console.WriteLine("[*] Word market readiness proof gallery");

            string galleryPath = Path.Combine(folderPath, "WordMarketReadinessProof");
            Directory.CreateDirectory(galleryPath);

            CreateTemplateAssemblyProof(galleryPath);
            CreateReviewDiffProof(galleryPath);
            CreateHtmlProof(galleryPath);
            CreateMarkdownProof(galleryPath);
            CreatePremiumWorkflowReportsProof(galleryPath);
            WriteGalleryIndex(galleryPath);
            WriteProofManifest(galleryPath);

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

            CreateInvoiceRowProof(scenarioPath);
            CreateContentControlFormProof(scenarioPath);
            WriteValidationReport(scenarioPath);
        }

        private static void CreateInvoiceRowProof(string scenarioPath) {
            string templatePath = Path.Combine(scenarioPath, "invoice-lines-template.docx");
            using (WordDocument template = WordDocument.Create(templatePath)) {
                template.AddParagraph("Invoice").Style = WordParagraphStyles.Heading1;
                template.AddParagraph("Customer: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Customer\"" });
                template.AddParagraph("Invoice number: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"InvoiceNumber\"" });

                WordTable table = template.AddTable(2, 3);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Item";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Qty";
                table.Rows[0].Cells[2].Paragraphs[0].Text = "Price";
                table.Rows[1].Cells[0].Paragraphs[0].AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Item\"" });
                table.Rows[1].Cells[1].Paragraphs[0].AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Quantity\"" });
                table.Rows[1].Cells[2].Paragraphs[0].AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Price\"" });

                WordMailMergeTemplateInspection inspection = WordMailMerge.InspectTemplate(
                    template,
                    new[] { "Customer", "InvoiceNumber", "Item", "Quantity", "Price" });
                WriteTemplateDiagnostics(Path.Combine(scenarioPath, "invoice-template-diagnostics.md"), inspection);
                template.Save();
            }

            var lineItems = new[] {
                new Dictionary<string, string> {
                    ["Item"] = "Assessment workshop",
                    ["Quantity"] = "1",
                    ["Price"] = "1200"
                },
                new Dictionary<string, string> {
                    ["Item"] = "Implementation sprint",
                    ["Quantity"] = "2",
                    ["Price"] = "3400"
                },
                new Dictionary<string, string> {
                    ["Item"] = "Readiness review",
                    ["Quantity"] = "1",
                    ["Price"] = "900"
                }
            };
            WriteRecords(Path.Combine(scenarioPath, "invoice-line-data.md"), lineItems);

            string outputPath = Path.Combine(scenarioPath, "invoice-lines-generated.docx");
            File.Copy(templatePath, outputPath, overwrite: true);
            using WordDocument invoice = WordDocument.Load(outputPath);
            WordMailMerge.ExecuteTableRows(invoice.Tables[0], templateRowIndex: 1, lineItems);
            WordMailMerge.Execute(invoice, new Dictionary<string, string> {
                ["Customer"] = "Northwind Traders",
                ["InvoiceNumber"] = "INV-2026-0042"
            });
            invoice.Save(false);
        }

        private static void CreateContentControlFormProof(string scenarioPath) {
            string filePath = Path.Combine(scenarioPath, "client-intake-form.docx");
            string logoSourcePath = Path.Combine(AppContext.BaseDirectory, "Images", "Kulek.jpg");
            string logoReplacementPath = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");

            using WordDocument document = WordDocument.Create(filePath);
            document.AddParagraph("Client Intake").Style = WordParagraphStyles.Heading1;
            document.AddStructuredDocumentTag("Client placeholder", "Client Alias", "ClientName");
            document.AddParagraph("Accepted:").AddCheckBox(false, "Accepted Alias", "Accepted");
            document.AddParagraph("Due:").AddDatePicker(new DateTime(2026, 1, 1), "Due Alias", "DueDate");
            document.AddParagraph("Priority:").AddDropDownList(new[] { "Low", "Medium", "High" }, "Priority Alias", "Priority");
            document.AddParagraph("Logo:").AddPictureControl(logoSourcePath, 24, 24, "Logo Alias", "Logo");
            document.AddParagraph("Tasks:").AddRepeatingSection("Tasks", "Tasks Alias", "Tasks");

            var values = new Dictionary<string, object?> {
                ["ClientName"] = "Northwind Traders",
                ["Accepted"] = true,
                ["DueDate"] = new DateTime(2026, 5, 29),
                ["Priority"] = "High",
                ["Logo"] = WordContentControlPictureValue.FromFile(logoReplacementPath),
                ["Tasks"] = new[] { "Collect source template", "Generate document", "Validate output" }
            };

            WordContentControlFormValidationResult validation = document.ValidateContentControlValues(values);
            WriteFormDiagnostics(Path.Combine(scenarioPath, "client-intake-validation.md"), validation);
            validation.EnsureValid();

            int updated = document.FillContentControlValues(values);
            WriteExtractedValues(Path.Combine(scenarioPath, "client-intake-filled-values.md"), updated, document.ExtractContentControlValues());
            document.Save(false);
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
                using WordDocument document = scenario.Value.ToWordDocument(options);
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

        private static void WriteFormDiagnostics(string path, WordContentControlFormValidationResult validation) {
            var builder = new StringBuilder();
            builder.AppendLine("# Content-Control Form Validation");
            builder.AppendLine();
            builder.AppendLine("- Valid: " + validation.IsValid);
            builder.AppendLine("- Expected keys: " + string.Join(", ", validation.ExpectedKeys));
            builder.AppendLine("- Supplied keys: " + string.Join(", ", validation.SuppliedKeys));
            builder.AppendLine("- Issues: " + validation.Issues.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
            foreach (WordContentControlFormIssue issue in validation.Issues) {
                builder.AppendLine("  - " + issue.Kind + " " + issue.Key + ": " + issue.Message);
            }

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        private static void WriteExtractedValues(string path, int updated, IReadOnlyDictionary<string, object?> values) {
            var builder = new StringBuilder();
            builder.AppendLine("# Filled Content-Control Values");
            builder.AppendLine();
            builder.AppendLine("- Updated controls: " + updated.ToString(System.Globalization.CultureInfo.InvariantCulture));
            foreach (KeyValuePair<string, object?> value in values.OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase)) {
                builder.AppendLine("- " + value.Key + ": " + FormatValue(value.Value));
            }

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        private static string FormatValue(object? value) {
            return value switch {
                null => string.Empty,
                DateTime date => date.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture),
                WordContentControlPictureValue picture => picture.FileName,
                IEnumerable<string> textItems => string.Join(", ", textItems),
                _ => Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty
            };
        }

        private static void WriteGalleryIndex(string galleryPath) {
            var builder = new StringBuilder();
            builder.AppendLine("# Word Market Readiness Proof Gallery");
            builder.AppendLine();
            builder.AppendLine("This folder is generated by `dotnet run --project OfficeIMO.Examples -- --word-market-readiness`.");
            builder.AppendLine("It is intentionally artifact-first: each scenario keeps source input, generated `.docx` output, diagnostics, and Open XML validation results together.");
            builder.AppendLine();
            builder.AppendLine("| Scenario | What it proves | Key artifacts |");
            builder.AppendLine("| --- | --- | --- |");
            foreach (ProofScenarioInfo scenario in GetProofScenarios()) {
                string scenarioPath = Path.Combine(galleryPath, scenario.DirectoryName);
                string artifacts = string.Join("<br>", Directory.GetFiles(scenarioPath)
                    .Select(path => Path.GetFileName(path))
                    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase));
                builder.AppendLine("| " + scenario.Title + " | " + scenario.Description + " | " + artifacts + " |");
            }

            File.WriteAllText(Path.Combine(galleryPath, "README.md"), builder.ToString(), Encoding.UTF8);
        }

        private static void WriteProofManifest(string galleryPath) {
            var scenarios = GetProofScenarios()
                .Select(scenario => {
                    string scenarioPath = Path.Combine(galleryPath, scenario.DirectoryName);
                    return new {
                        scenario = scenario.DirectoryName,
                        title = scenario.Title,
                        description = scenario.Description,
                        artifacts = Directory.GetFiles(scenarioPath)
                            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                            .Select(path => new {
                                path = ToRelativeGalleryPath(galleryPath, path),
                                kind = Path.GetExtension(path).TrimStart('.').ToLowerInvariant(),
                                validated = string.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase),
                                validationErrors = string.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase)
                                    ? ValidateDocx(path).Count
                                    : -1
                            })
                            .ToArray()
                    };
                })
                .ToArray();

            var manifest = new {
                generatedBy = "OfficeIMO.Examples --word-market-readiness",
                scope = "OfficeIMO.Word non-PDF market readiness",
                scenarios
            };

            File.WriteAllText(
                Path.Combine(galleryPath, "proof-manifest.json"),
                JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }),
                Encoding.UTF8);
        }

        private static string ToRelativeGalleryPath(string galleryPath, string path) {
            return Path.GetRelativePath(galleryPath, path).Replace(Path.DirectorySeparatorChar, '/');
        }

        private static IReadOnlyList<ProofScenarioInfo> GetProofScenarios() {
            return new[] {
                new ProofScenarioInfo(
                    "01-template-assembly",
                    "Template assembly",
                    "Batch merge, repeated table rows, content-control form fill, diagnostics, and generated documents."),
                new ProofScenarioInfo(
                    "02-review-diff",
                    "Review and structured diff",
                    "Document comparison with paragraph, table, row, cell, and image findings."),
                new ProofScenarioInfo(
                    "03-html-conversion",
                    "HTML conversion",
                    "Untrusted HTML import, generated DOCX, round-trip HTML, diagnostics, and validation."),
                new ProofScenarioInfo(
                    "04-markdown-conversion",
                    "Markdown conversion",
                    "Markdown import, generated DOCX, round-trip Markdown, diagnostics, and validation."),
                new ProofScenarioInfo(
                    "05-premium-workflow-reports",
                    "Premium workflow reports",
                    "Unknown-document feature preflight, review reports, comparison reports, generated redline artifact, field refresh diagnostics, template preflight, and signature preflight.")
            };
        }

        private static IReadOnlyList<ValidationErrorInfo> ValidateDocx(string path) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, false);
            var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
            return validator.Validate(document).ToList();
        }

        private sealed class ProofScenarioInfo {
            internal ProofScenarioInfo(string directoryName, string title, string description) {
                DirectoryName = directoryName;
                Title = title;
                Description = description;
            }

            internal string DirectoryName { get; }

            internal string Title { get; }

            internal string Description { get; }
        }
    }
}
