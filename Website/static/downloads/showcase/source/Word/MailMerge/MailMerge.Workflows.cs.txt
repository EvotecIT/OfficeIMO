using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class MailMerge {
        internal static void Example_MailMergeInvoiceWorkflow(string folderPath, bool openWord) {
            Console.WriteLine("[*] Mail merge invoice workflow");
            string filePath = Path.Combine(folderPath, "MailMergeInvoiceWorkflow.docx");
            using WordDocument document = WordDocument.Create(filePath);

            document.AddParagraph("Invoice").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("Customer: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Customer\"" });
            document.AddParagraph("Invoice number: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"InvoiceNumber\"" });
            document.AddParagraph("Issued: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"IssuedDate\"" });

            WordTable table = document.AddTable(2, 4);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Item";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Quantity";
            table.Rows[0].Cells[2].Paragraphs[0].Text = "Unit price";
            table.Rows[0].Cells[3].Paragraphs[0].Text = "Line total";
            table.Rows[1].Cells[0].Paragraphs[0].AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Item\"" });
            table.Rows[1].Cells[1].Paragraphs[0].AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Quantity\"" });
            table.Rows[1].Cells[2].Paragraphs[0].AddField(WordFieldType.MergeField, parameters: new List<string> { "\"UnitPrice\"" });
            table.Rows[1].Cells[3].Paragraphs[0].AddField(WordFieldType.MergeField, parameters: new List<string> { "\"LineTotal\"" });

            WordTemplatePreflightReport preflight = WordMailMerge.PreflightTemplate(
                document,
                mergeFieldNames: new[] { "Customer", "InvoiceNumber", "IssuedDate", "Item", "Quantity", "UnitPrice", "LineTotal" });
            WriteTemplatePreflight(Path.Combine(folderPath, "MailMergeInvoiceWorkflow.Preflight.md"), preflight);
            preflight.EnsureCan(WordTemplatePreflightCapability.BindTemplate);

            var rows = new[] {
                new Dictionary<string, string> {
                    ["Item"] = "Assessment workshop",
                    ["Quantity"] = "1",
                    ["UnitPrice"] = "1200",
                    ["LineTotal"] = "1200"
                },
                new Dictionary<string, string> {
                    ["Item"] = "Implementation sprint",
                    ["Quantity"] = "2",
                    ["UnitPrice"] = "3400",
                    ["LineTotal"] = "6800"
                },
                new Dictionary<string, string> {
                    ["Item"] = "Readiness review",
                    ["Quantity"] = "1",
                    ["UnitPrice"] = "900",
                    ["LineTotal"] = "900"
                }
            };

            WordMailMerge.ExecuteTableRows(table, templateRowIndex: 1, rows);
            WordMailMerge.Execute(document, new Dictionary<string, string> {
                ["Customer"] = "Northwind Traders",
                ["InvoiceNumber"] = "INV-2026-0042",
                ["IssuedDate"] = "2026-06-28"
            });
            document.AddParagraph("Total: 8900");
            document.Save();
            if (openWord) document.OpenInApplication();
        }

        internal static void Example_MailMergeGroupedTableWorkflow(string folderPath, bool openWord) {
            Console.WriteLine("[*] Mail merge grouped table workflow");
            string filePath = Path.Combine(folderPath, "MailMergeGroupedTableWorkflow.docx");
            using WordDocument document = WordDocument.Create(filePath);

            document.AddParagraph("Grouped Revenue Report").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("Client: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"ClientName\"" });
            document.AddParagraph("Period: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Period\"" });

            WordTable table = document.AddTable(3, 3);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Category";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Item";
            table.Rows[0].Cells[2].Paragraphs[0].Text = "Amount";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Group: ";
            table.Rows[1].Cells[0].Paragraphs[0].AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Category\"" });
            table.Rows[1].Cells[1].Paragraphs[0].Text = "Subtotal";
            table.Rows[1].Cells[2].Paragraphs[0].AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Subtotal\"" });
            table.Rows[2].Cells[0].Paragraphs[0].Text = string.Empty;
            table.Rows[2].Cells[1].Paragraphs[0].AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Item\"" });
            table.Rows[2].Cells[2].Paragraphs[0].AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Amount\"" });

            WordTemplatePreflightReport preflight = WordMailMerge.PreflightTemplate(
                document,
                mergeFieldNames: new[] { "ClientName", "Period", "Category", "Subtotal", "Item", "Amount" });
            WriteTemplatePreflight(Path.Combine(folderPath, "MailMergeGroupedTableWorkflow.Preflight.md"), preflight);
            preflight.EnsureCan(WordTemplatePreflightCapability.BindTemplate);

            WordMailMergeTableRowGroupResult result = WordMailMerge.ExecuteTableRowGroups(
                table,
                groupTemplateRowIndex: 1,
                detailTemplateRowIndex: 2,
                groups: new[] {
                    new WordMailMergeTableRowGroup(
                        new Dictionary<string, string> {
                            ["Category"] = "Consulting",
                            ["Subtotal"] = "5200"
                        },
                        new[] {
                            new Dictionary<string, string> {
                                ["Item"] = "Discovery workshop",
                                ["Amount"] = "1800"
                            },
                            new Dictionary<string, string> {
                                ["Item"] = "Implementation planning",
                                ["Amount"] = "3400"
                            }
                        }),
                    new WordMailMergeTableRowGroup(
                        new Dictionary<string, string> {
                            ["Category"] = "Managed service",
                            ["Subtotal"] = "2700"
                        },
                        new[] {
                            new Dictionary<string, string> {
                                ["Item"] = "Monthly operations",
                                ["Amount"] = "2100"
                            },
                            new Dictionary<string, string> {
                                ["Item"] = "Readiness review",
                                ["Amount"] = "600"
                            }
                        })
                });

            WordMailMerge.Execute(document, new Dictionary<string, string> {
                ["ClientName"] = "Fabrikam Finance",
                ["Period"] = "Q3 2026"
            });
            document.AddParagraph("Generated rows: " + result.TotalRowCount.ToString(System.Globalization.CultureInfo.InvariantCulture));
            document.Save();
            if (openWord) document.OpenInApplication();
        }

        internal static void Example_MailMergeProposalWorkflow(string folderPath, bool openWord) {
            Console.WriteLine("[*] Mail merge proposal workflow");
            string filePath = Path.Combine(folderPath, "MailMergeProposalWorkflow.docx");
            using WordDocument document = WordDocument.Create(filePath);

            document.AddParagraph("Proposal").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("Client: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Client\"" });
            document.AddParagraph("Prepared by: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"PreparedBy\"" });
            document.AddParagraph("{{#ShowExecutiveSummary}}");
            document.AddParagraph("Executive summary: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"ExecutiveSummary\"" });
            document.AddParagraph("{{/ShowExecutiveSummary}}");
            document.AddParagraph("{{#each Services}}");
            document.AddParagraph("Service: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"ServiceName\"" });
            document.AddParagraph("{{#each Deliverables}}");
            document.AddParagraph("Deliverable: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Deliverable\"" });
            document.AddParagraph("{{/each Deliverables}}");
            document.AddParagraph("{{/each Services}}");

            WordTemplatePreflightReport preflight = WordMailMerge.PreflightTemplate(
                document,
                mergeFieldNames: new[] { "Client", "PreparedBy", "ExecutiveSummary", "ServiceName", "Deliverable" },
                conditionNames: new[] { "ShowExecutiveSummary" },
                repeatingBlockNames: new[] { "Services", "Deliverables" });
            WriteTemplatePreflight(Path.Combine(folderPath, "MailMergeProposalWorkflow.Preflight.md"), preflight);
            preflight.EnsureCan(WordTemplatePreflightCapability.BindTemplate);

            WordMailMerge.ExecuteConditionalBlocks(document, new Dictionary<string, bool> {
                ["ShowExecutiveSummary"] = true
            });
            WordMailMerge.ExecuteRepeatingBlockRegions(
                document,
                new Dictionary<string, IEnumerable<WordMailMergeBlockData>> {
                    ["Services"] = new[] {
                        new WordMailMergeBlockData(
                            new Dictionary<string, string> {
                                ["ServiceName"] = "Document automation readiness"
                            },
                            new Dictionary<string, IEnumerable<WordMailMergeBlockData>> {
                                ["Deliverables"] = new[] {
                                    new WordMailMergeBlockData(new Dictionary<string, string> {
                                        ["Deliverable"] = "Template inventory"
                                    }),
                                    new WordMailMergeBlockData(new Dictionary<string, string> {
                                        ["Deliverable"] = "Generated proof package"
                                    })
                                }
                            }),
                        new WordMailMergeBlockData(
                            new Dictionary<string, string> {
                                ["ServiceName"] = "Review workflow rollout"
                            },
                            new Dictionary<string, IEnumerable<WordMailMergeBlockData>> {
                                ["Deliverables"] = new[] {
                                    new WordMailMergeBlockData(new Dictionary<string, string> {
                                        ["Deliverable"] = "Review report"
                                    }),
                                    new WordMailMergeBlockData(new Dictionary<string, string> {
                                        ["Deliverable"] = "Redline artifact"
                                    })
                                }
                            })
                    }
                });
            WordMailMerge.Execute(document, new Dictionary<string, string> {
                ["Client"] = "Contoso Legal",
                ["PreparedBy"] = "OfficeIMO",
                ["ExecutiveSummary"] = "This proposal is generated from a reusable Word template with conditional and nested repeated regions."
            });
            document.Save();
            if (openWord) document.OpenInApplication();
        }

        internal static void Example_MailMergeReviewLetterWorkflow(string folderPath, bool openWord) {
            Console.WriteLine("[*] Mail merge review letter workflow");
            string filePath = Path.Combine(folderPath, "MailMergeReviewLetterWorkflow.docx");
            using WordDocument document = WordDocument.Create(filePath);

            document.AddParagraph("Review Letter").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("To: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Recipient\"" });
            document.AddParagraph("Matter: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Matter\"" });
            document.AddParagraph("{{#IncludeFindings}}");
            document.AddParagraph("Open findings").Style = WordParagraphStyles.Heading2;
            document.AddParagraph("{{#each Findings}}");
            document.AddParagraph("Finding: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"FindingTitle\"" });
            document.AddParagraph("Owner: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Owner\"" });
            document.AddParagraph("Due: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"DueDate\"" });
            document.AddParagraph("{{/each Findings}}");
            document.AddParagraph("{{/IncludeFindings}}");
            WordParagraph closing = document.AddParagraph("Please confirm the remediation dates by close of business.");

            WordTemplatePreflightReport preflight = WordMailMerge.PreflightTemplate(
                document,
                mergeFieldNames: new[] { "Recipient", "Matter", "FindingTitle", "Owner", "DueDate" },
                conditionNames: new[] { "IncludeFindings" },
                repeatingBlockNames: new[] { "Findings" });
            WriteTemplatePreflight(Path.Combine(folderPath, "MailMergeReviewLetterWorkflow.Preflight.md"), preflight);
            preflight.EnsureCan(WordTemplatePreflightCapability.BindTemplate);

            WordMailMerge.ExecuteConditionalBlocks(document, new Dictionary<string, bool> {
                ["IncludeFindings"] = true
            });
            WordMailMerge.ExecuteRepeatingBlocks(
                document,
                new Dictionary<string, IEnumerable<IDictionary<string, string>>> {
                    ["Findings"] = new[] {
                        new Dictionary<string, string> {
                            ["FindingTitle"] = "Update retention language",
                            ["Owner"] = "Legal Operations",
                            ["DueDate"] = "2026-07-10"
                        },
                        new Dictionary<string, string> {
                            ["FindingTitle"] = "Confirm incident notice period",
                            ["Owner"] = "Security",
                            ["DueDate"] = "2026-07-17"
                        }
                    }
                });
            WordMailMerge.Execute(document, new Dictionary<string, string> {
                ["Recipient"] = "Document Owner",
                ["Matter"] = "Services agreement review"
            });
            closing.AddComment("Reviewer", "RV", "This generated letter keeps review feedback close to the produced document.");
            document.Save();
            if (openWord) document.OpenInApplication();
        }

        internal static void Example_MailMergeHeaderFooterWorkflow(string folderPath, bool openWord) {
            Console.WriteLine("[*] Mail merge header/footer workflow");
            string filePath = Path.Combine(folderPath, "MailMergeHeaderFooterWorkflow.docx");
            using WordDocument document = WordDocument.Create(filePath);

            document.AddParagraph("Approval Package").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("Package: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"PackageTitle\"" });
            document.AddParagraph("This workflow proves template markers in headers and footers through the same preflight and binding APIs used for body content.");
            document.AddHeadersAndFooters();

            WordHeader header = document.Header?.Default
                ?? throw new InvalidOperationException("The default header was not created.");
            header.AddParagraph("{{#ShowConfidentialHeader}}");
            header.AddParagraph("Client: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"ClientName\"" });
            header.AddParagraph("Prepared: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"PreparedDate\"" });
            header.AddParagraph("{{/ShowConfidentialHeader}}");

            WordFooter footer = document.Footer?.Default
                ?? throw new InvalidOperationException("The default footer was not created.");
            footer.AddParagraph("{{#each Signers}}");
            footer.AddParagraph("Signer: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"SignerName\"" });
            footer.AddParagraph("Role: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"SignerRole\"" });
            footer.AddParagraph("{{/each Signers}}");

            WordTemplatePreflightReport preflight = WordMailMerge.PreflightTemplate(
                document,
                mergeFieldNames: new[] { "PackageTitle", "ClientName", "PreparedDate", "SignerName", "SignerRole" },
                conditionNames: new[] { "ShowConfidentialHeader" },
                repeatingBlockNames: new[] { "Signers" });
            WriteTemplatePreflight(Path.Combine(folderPath, "MailMergeHeaderFooterWorkflow.Preflight.md"), preflight);
            preflight.EnsureCan(WordTemplatePreflightCapability.BindTemplate);

            WordMailMerge.ExecuteConditionalBlocks(document, new Dictionary<string, bool> {
                ["ShowConfidentialHeader"] = true
            });
            WordMailMerge.ExecuteRepeatingBlocks(
                document,
                new Dictionary<string, IEnumerable<IDictionary<string, string>>> {
                    ["Signers"] = new[] {
                        new Dictionary<string, string> {
                            ["SignerName"] = "Avery Stone",
                            ["SignerRole"] = "Business Owner"
                        },
                        new Dictionary<string, string> {
                            ["SignerName"] = "Morgan Lee",
                            ["SignerRole"] = "Security Reviewer"
                        }
                    }
                });
            WordMailMerge.Execute(document, new Dictionary<string, string> {
                ["PackageTitle"] = "Contract automation rollout",
                ["ClientName"] = "Contoso Legal",
                ["PreparedDate"] = "2026-06-30"
            });
            document.Save();
            if (openWord) document.OpenInApplication();
        }

        internal static void Example_MailMergeFormFillWorkflow(string folderPath, bool openWord) {
            Console.WriteLine("[*] Mail merge form-fill workflow");
            string filePath = Path.Combine(folderPath, "MailMergeFormFillWorkflow.docx");
            using WordDocument document = WordDocument.Create(filePath);

            document.AddParagraph("Client Intake Form").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("Client: ").AddStructuredDocumentTag("Client placeholder", "Client", "ClientName");
            document.AddParagraph("Accepted: ").AddCheckBox(false, "Accepted", "Accepted");
            document.AddParagraph("Due date: ").AddDatePicker(new DateTime(2026, 1, 1), "Due date", "DueDate");
            document.AddParagraph("Priority: ").AddDropDownList(new[] { "Low", "Medium", "High" }, "Priority", "Priority");
            document.AddParagraph("Tasks: ").AddRepeatingSection("Tasks", "Tasks", "Tasks");

            var values = new Dictionary<string, object?> {
                ["ClientName"] = "Northwind Traders",
                ["Accepted"] = true,
                ["DueDate"] = new DateTime(2026, 7, 15),
                ["Priority"] = "High",
                ["Tasks"] = new[] { "Collect template", "Bind form values", "Validate generated document" }
            };

            WordContentControlFormValidationResult validation = document.ValidateContentControlValues(values);
            WriteFormValidation(
                Path.Combine(folderPath, "MailMergeFormFillWorkflow.Validation.md"),
                Path.Combine(folderPath, "MailMergeFormFillWorkflow.Validation.json"),
                validation);
            validation.EnsureValid();

            int updated = document.FillContentControlValues(values);
            WriteExtractedValues(Path.Combine(folderPath, "MailMergeFormFillWorkflow.ExtractedValues.md"), updated, document.ExtractContentControlValues());
            document.Save();
            if (openWord) document.OpenInApplication();
        }

        internal static void Example_MailMergeWorkflowGallery(string folderPath, bool openWord) {
            Example_MailMergeInvoiceWorkflow(folderPath, openWord);
            Example_MailMergeGroupedTableWorkflow(folderPath, openWord);
            Example_MailMergeProposalWorkflow(folderPath, openWord);
            Example_MailMergeReviewLetterWorkflow(folderPath, openWord);
            Example_MailMergeHeaderFooterWorkflow(folderPath, openWord);
            Example_MailMergeFormFillWorkflow(folderPath, openWord);
        }

        private static void WriteTemplatePreflight(string path, WordTemplatePreflightReport report) {
            File.WriteAllText(path, report.ToMarkdown(), Encoding.UTF8);
        }

        private static void WriteFormValidation(string markdownPath, string jsonPath, WordContentControlFormValidationResult validation) {
            File.WriteAllText(markdownPath, validation.ToMarkdown(), Encoding.UTF8);
            File.WriteAllText(jsonPath, validation.ToJson(), Encoding.UTF8);
        }

        private static void WriteExtractedValues(string path, int updated, IReadOnlyDictionary<string, object?> values) {
            var builder = new StringBuilder();
            builder.AppendLine("# Extracted Content-Control Values");
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
                IEnumerable<string> items => string.Join(", ", items),
                _ => Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty
            };
        }
    }
}
