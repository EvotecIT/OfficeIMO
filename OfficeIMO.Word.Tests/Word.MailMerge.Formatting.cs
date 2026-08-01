using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_MailMerge_BatchReportKeepsPerRecordMissingValueDiagnostics() {
            string templatePath = Path.Combine(_directoryWithFiles, "MailMergeBatchReportTemplate.docx");
            using (WordDocument template = WordDocument.Create(templatePath)) {
                template.AddParagraph("Name: ").AddField(WordFieldType.MergeField, parameters: new List<string> { "Name" });
                template.AddParagraph("City: ").AddField(WordFieldType.MergeField, parameters: new List<string> { "City" });
                template.Save();
            }

            var records = new[] {
                (IDictionary<string, string>)new Dictionary<string, string> { ["Name"] = "Ada", ["City"] = "London" },
                new Dictionary<string, string> { ["Name"] = "Grace" }
            };

            WordMailMergeBatchResult result = WordMailMerge.ExecuteBatchWithReport(
                templatePath,
                records,
                (index, _) => Path.Combine(_directoryWithFiles, $"MailMergeBatchReport-{index}.docx"));

            Assert.Equal(2, result.Items.Count);
            Assert.True(result.Items[0].Execution.IsComplete);
            Assert.False(result.Items[1].Execution.IsComplete);
            Assert.Equal(new[] { "City" }, result.Items[1].Execution.MissingValueNames);
            Assert.False(result.IsComplete);
            Assert.Throws<InvalidOperationException>(() => result.EnsureComplete());
        }

        [Fact]
        public void Test_MailMerge_ComplexSplitRunFieldsPreserveResultFormattingWhenKeepingFields() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeFormattingSplitComplexFields.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Body body = document._document.MainDocumentPart!.Document.Body!;
                body.Append(new Paragraph(
                    new Run(new Text("Client: ")),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" MERGE")),
                    new Run(new FieldCode("FIELD \"Client\" ")),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                    new Run(
                        new RunProperties(new Bold(), new Color { Val = "C00000" }),
                        new Text("Place")),
                    new Run(
                        new RunProperties(new Italic(), new Color { Val = "008000" }),
                        new Text("holder")),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End })));

                WordMailMerge.Execute(
                    document,
                    new Dictionary<string, string> {
                        ["Client"] = "Northwind Traders"
                    },
                    removeFields: false);
                document.Save();
            }

            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false);
            Body bodyXml = wordDocument.MainDocumentPart!.Document.Body!;
            Assert.Contains("MERGE", bodyXml.InnerXml);
            Assert.Contains("FIELD", bodyXml.InnerXml);
            Assert.Contains("Northwind Traders", bodyXml.InnerText);
            Assert.DoesNotContain("Placeholder", bodyXml.InnerText);

            Run replacementRun = Assert.Single(bodyXml.Descendants<Run>(), run => run.Elements<Text>().Any(text => text.Text == "Northwind Traders"));
            Assert.NotNull(replacementRun.RunProperties?.Bold);
            Assert.Equal("C00000", replacementRun.RunProperties!.Color!.Val!.Value);

            Run emptiedRun = Assert.Single(bodyXml.Descendants<Run>(), run => run.Elements<Text>().Any(text => text.Text == string.Empty));
            Assert.NotNull(emptiedRun.RunProperties?.Italic);
            Assert.Equal("008000", emptiedRun.RunProperties!.Color!.Val!.Value);
        }

        [Fact]
        public void Test_MailMerge_KeepingFieldsCreatesMissingSimpleAndComplexResultRuns() {
            using WordDocument document = WordDocument.Create();
            Body body = document._document.MainDocumentPart!.Document.Body!;
            body.Append(
                new Paragraph(new SimpleField { Instruction = " MERGEFIELD SimpleName " }),
                new Paragraph(
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" MERGEFIELD ComplexName ")),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End })));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> {
                    ["SimpleName"] = "Simple value",
                    ["ComplexName"] = "Complex value"
                },
                removeFields: false);

            Assert.True(report.IsComplete);
            Assert.Equal(2, report.MergedCount);
            Assert.Contains("Simple value", body.InnerText);
            Assert.Contains("Complex value", body.InnerText);
            Assert.Contains("MERGEFIELD SimpleName", body.InnerXml);
            Assert.Contains("MERGEFIELD ComplexName", body.InnerXml);
        }

        [Fact]
        public void Test_MailMerge_ReportsAndUpdatesComplexFieldInsideInlineContentControl() {
            using WordDocument document = WordDocument.Create();
            Body body = document._document.MainDocumentPart!.Document.Body!;
            body.Append(new Paragraph(new SdtRun(
                new SdtProperties(new Tag { Val = "ClientName" }),
                new SdtContentRun(
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" MERGEFIELD ClientName ")),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("stale")),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End })))));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["ClientName"] = "Northwind" },
                removeFields: false);

            Assert.True(report.IsComplete);
            Assert.Equal(1, report.MergedCount);
            Assert.Contains("Northwind", body.InnerText);
            Assert.DoesNotContain("stale", body.InnerText);
            Assert.Contains("MERGEFIELD ClientName", body.InnerXml);
        }

        [Fact]
        public void Test_MailMerge_NestedRegionsPreserveTableCellFieldFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeFormattingNestedRegionsTableCells.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("{{#each Projects}}");
                WordTable table = document.AddTable(1, 2);
                ReplaceCellContentForFormattingTest(
                    table.Rows[0].Cells[0]._tableCell,
                    new Paragraph(
                        new Run(new Text("Project: ")),
                        CreateSimpleMergeFieldForFormattingTest("ProjectName", new RunProperties(new Bold(), new Color { Val = "1F4E79" }))));
                ReplaceCellContentForFormattingTest(
                    table.Rows[0].Cells[1]._tableCell,
                    new Paragraph(new Run(new Text("{{#each Tasks}}"))),
                    new Paragraph(
                        new Run(new Text("Task: ")),
                        CreateSimpleMergeFieldForFormattingTest("TaskName", new RunProperties(new Italic(), new Color { Val = "008000" }))),
                    new Paragraph(new Run(new Text("{{/each Tasks}}"))));
                document.AddParagraph("{{/each Projects}}");

                int generated = WordMailMerge.ExecuteRepeatingBlockRegions(
                    document,
                    new Dictionary<string, IEnumerable<WordMailMergeBlockData>> {
                        ["Projects"] = new[] {
                            new WordMailMergeBlockData(
                                new Dictionary<string, string> {
                                    ["ProjectName"] = "Readiness"
                                },
                                new Dictionary<string, IEnumerable<WordMailMergeBlockData>> {
                                    ["Tasks"] = new[] {
                                        new WordMailMergeBlockData(new Dictionary<string, string> {
                                            ["TaskName"] = "Design"
                                        }),
                                        new WordMailMergeBlockData(new Dictionary<string, string> {
                                            ["TaskName"] = "Validate"
                                        })
                                    }
                                }),
                            new WordMailMergeBlockData(
                                new Dictionary<string, string> {
                                    ["ProjectName"] = "Rollout"
                                },
                                new Dictionary<string, IEnumerable<WordMailMergeBlockData>> {
                                    ["Tasks"] = new[] {
                                        new WordMailMergeBlockData(new Dictionary<string, string> {
                                            ["TaskName"] = "Publish"
                                        })
                                    }
                                })
                        }
                    });

                Assert.Equal(5, generated);
                document.Save();
            }

            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false);
            Body body = wordDocument.MainDocumentPart!.Document.Body!;
            Assert.Contains("Project: Readiness", body.InnerText);
            Assert.Contains("Task: Design", body.InnerText);
            Assert.Contains("Task: Validate", body.InnerText);
            Assert.Contains("Project: Rollout", body.InnerText);
            Assert.Contains("Task: Publish", body.InnerText);
            Assert.DoesNotContain("{{#each Projects}}", body.InnerText);
            Assert.DoesNotContain("{{#each Tasks}}", body.InnerText);
            Assert.DoesNotContain("MERGEFIELD", body.InnerXml);

            Run projectRun = Assert.Single(body.Descendants<Run>(), run => run.Elements<Text>().Any(text => text.Text == "Readiness"));
            Assert.NotNull(projectRun.RunProperties?.Bold);
            Assert.Equal("1F4E79", projectRun.RunProperties!.Color!.Val!.Value);

            Run taskRun = Assert.Single(body.Descendants<Run>(), run => run.Elements<Text>().Any(text => text.Text == "Design"));
            Assert.NotNull(taskRun.RunProperties?.Italic);
            Assert.Equal("008000", taskRun.RunProperties!.Color!.Val!.Value);
        }

        [Fact]
        public void Test_MailMerge_ContentControlFormFillPreservesTextRunFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeFormattingContentControlForm.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordStructuredDocumentTag tag = document.AddParagraph("Client: ")
                    .AddStructuredDocumentTag("Placeholder", "Client Alias", "ClientName");
                tag.Bold = true;
                tag.Color = OfficeIMO.Drawing.OfficeColor.ParseHex("#7030A0");

                int updated = document.FillContentControlValues(new Dictionary<string, object?> {
                    ["ClientName"] = "Northwind Traders"
                });

                Assert.Equal(1, updated);
                document.Save();
            }

            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false);
            Body body = wordDocument.MainDocumentPart!.Document.Body!;
            SdtRun sdtRun = Assert.Single(body.Descendants<SdtRun>());
            Run run = Assert.Single(sdtRun.SdtContentRun!.Elements<Run>());
            Text text = Assert.Single(run.Elements<Text>());
            Assert.Equal("Northwind Traders", text.Text);
            Assert.NotNull(run.RunProperties?.Bold);
            Assert.Equal("7030A0", run.RunProperties!.Color!.Val!.Value.ToUpperInvariant());
        }

        [Fact]
        public void Test_MailMerge_ExecutionReportFormatsSupportedPicturesAndReportsMissingValues() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeExecutionReport.docx");
            using WordDocument document = WordDocument.Create(filePath);
            Body body = document._document.MainDocumentPart!.Document.Body!;
            body.Append(
                new Paragraph(new SimpleField(new Run(new Text("name"))) { Instruction = @" MERGEFIELD Name \* Upper " }),
                new Paragraph(new SimpleField(new Run(new Text("amount"))) { Instruction = " MERGEFIELD Amount \\# \"#,##0.00\" " }),
                new Paragraph(new SimpleField(new Run(new Text("date"))) { Instruction = " MERGEFIELD DueDate \\@ \"yyyy-MM-dd\" " }),
                new Paragraph(new SimpleField(new Run(new Text("missing"))) { Instruction = " MERGEFIELD Missing " }));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> {
                    ["Name"] = "Northwind traders",
                    ["Amount"] = "1234.5",
                    ["DueDate"] = "2026-07-31T12:30:00Z"
                });

            Assert.Equal(3, report.MergedCount);
            Assert.False(report.IsComplete);
            Assert.Equal("Missing", Assert.Single(report.MissingValueNames));
            Assert.Contains(report.Fields, result => result.Name == "Name" && result.Value == "NORTHWIND TRADERS");
            Assert.Contains(report.Fields, result => result.Name == "Amount" && result.Value == "1,234.50");
            Assert.Contains(report.Fields, result => result.Name == "DueDate" && result.Value == "2026-07-31");
            Assert.Throws<System.InvalidOperationException>(() => report.EnsureComplete());
            Assert.Contains("NORTHWIND TRADERS", body.InnerText, System.StringComparison.Ordinal);
            Assert.Contains("1,234.50", body.InnerText, System.StringComparison.Ordinal);
            Assert.Contains("2026-07-31", body.InnerText, System.StringComparison.Ordinal);
            Assert.Contains("missing", body.InnerText, System.StringComparison.Ordinal);
        }

        [Fact]
        public void Test_MailMerge_DatePicturePreservesExplicitOffsetWallClock() {
            using WordDocument document = WordDocument.Create();
            Body body = document._document.MainDocumentPart!.Document.Body!;
            body.Append(new Paragraph(new SimpleField(new Run(new Text("date"))) {
                Instruction = " MERGEFIELD EventTime \\@ \"yyyy-MM-dd HH:mm zzz\" "
            }));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> {
                    ["EventTime"] = "2026-07-31T10:00:00+02:00"
                });

            WordMailMergeFieldResult result = Assert.Single(report.Fields);
            Assert.Equal(WordMailMergeFieldStatus.Merged, result.Status);
            Assert.Equal("2026-07-31 10:00 +02:00", result.Value);
            Assert.Contains("2026-07-31 10:00 +02:00", body.InnerText, System.StringComparison.Ordinal);
        }

        [Fact]
        public void Test_MailMerge_ExecutionReportIncludesMalformedSimpleAndComplexFields() {
            using WordDocument document = WordDocument.Create();
            Body body = document._document.MainDocumentPart!.Document.Body!;
            body.Append(
                new Paragraph(new SimpleField(new Run(new Text("malformed"))) { Instruction = " MERGEFIELD " }),
                new Paragraph(
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" MERGEFIELD "))));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string>());

            Assert.False(report.IsComplete);
            Assert.Equal(2, report.Fields.Count(result => result.Status == WordMailMergeFieldStatus.MalformedField));
            Assert.Throws<System.InvalidOperationException>(() => report.EnsureComplete());
        }

        [Fact]
        public void Test_MailMerge_UnsupportedSwitchLeavesFieldInPlaceAndReportsIt() {
            using WordDocument document = WordDocument.Create();
            Body body = document._document.MainDocumentPart!.Document.Body!;
            body.Append(new Paragraph(new SimpleField(new Run(new Text("placeholder"))) {
                Instruction = " MERGEFIELD Name \\b \"Dear \" "
            }));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["Name"] = "Ada" });

            WordMailMergeFieldResult result = Assert.Single(report.Fields);
            Assert.Equal(WordMailMergeFieldStatus.UnsupportedFormatting, result.Status);
            Assert.Contains("\\b", result.Message, System.StringComparison.Ordinal);
            Assert.Single(body.Descendants<SimpleField>());
            Assert.Contains("placeholder", body.InnerText, System.StringComparison.Ordinal);
            Assert.DoesNotContain("Ada", body.InnerText, System.StringComparison.Ordinal);
        }

        [Fact]
        public void Test_MailMerge_ExecutionReportPreservesSimpleAndComplexOccurrenceOrder() {
            using WordDocument document = WordDocument.Create();
            Body body = document._document.MainDocumentPart!.Document.Body!;
            body.Append(new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" MERGEFIELD First ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("first placeholder")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new Text(" / ")),
                new SimpleField(new Run(new Text("second placeholder"))) { Instruction = " MERGEFIELD Second " }));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["First"] = "one", ["Second"] = "two" },
                removeFields: false);

            Assert.Equal(new[] { "First", "Second" }, report.Fields.Select(result => result.Name));
            Assert.Equal(2, report.MergedCount);
        }

        [Fact]
        public void Test_MailMerge_ExecutionReportDoesNotLoseOuterMergeFieldWhenFieldsAreNested() {
            using WordDocument document = WordDocument.Create();
            Body body = document._document.MainDocumentPart!.Document.Body!;
            body.Append(new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" MERGEFIELD Outer ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" MERGEFIELD Inner ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("inner")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End })));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["Outer"] = "outer value", ["Inner"] = "inner value" });

            Assert.False(report.IsComplete);
            Assert.Contains(report.Fields, result => result.Status == WordMailMergeFieldStatus.MalformedField);
            Assert.Contains(report.Fields, result => result.Name == "Inner" && result.Status == WordMailMergeFieldStatus.Merged);
        }

        [Fact]
        public void Test_MailMerge_ExecutionReportPreservesOuterComplexFieldContainingNestedSimpleField() {
            using WordDocument document = WordDocument.Create();
            Body body = document._document.MainDocumentPart!.Document.Body!;
            body.Append(new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" MERGEFIELD Outer ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new SimpleField(new Run(new Text("inner placeholder"))) { Instruction = " MERGEFIELD Inner " },
                new Run(new FieldChar { FieldCharType = FieldCharValues.End })));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["Outer"] = "outer value", ["Inner"] = "inner value" },
                removeFields: false);

            Assert.False(report.IsComplete);
            Assert.Contains(report.Fields, result =>
                result.Status == WordMailMergeFieldStatus.MalformedField &&
                result.Instruction.Contains("Outer", System.StringComparison.Ordinal));
            Assert.Contains(report.Fields, result =>
                result.Name == "Inner" &&
                result.Status == WordMailMergeFieldStatus.Merged &&
                result.Value == "inner value");
            Assert.DoesNotContain("outer value", body.InnerText, System.StringComparison.Ordinal);
            Assert.Contains("inner value", body.InnerText, System.StringComparison.Ordinal);
        }

        [Fact]
        public void Test_MailMerge_ExecutionReportRejectsOuterSimpleFieldContainingNestedSimpleField() {
            using WordDocument document = WordDocument.Create();
            Body body = document._document.MainDocumentPart!.Document.Body!;
            body.Append(new Paragraph(
                new SimpleField(
                    new Run(new Text("outer placeholder")),
                    new SimpleField(new Run(new Text("inner placeholder"))) { Instruction = " MERGEFIELD Inner " }) {
                    Instruction = " MERGEFIELD Outer "
                }));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["Outer"] = "outer value", ["Inner"] = "inner value" });

            Assert.False(report.IsComplete);
            Assert.Contains(report.Fields, result =>
                result.Status == WordMailMergeFieldStatus.MalformedField &&
                result.Instruction.Contains("Outer", System.StringComparison.Ordinal));
            Assert.Contains(report.Fields, result =>
                result.Name == "Inner" &&
                result.Status == WordMailMergeFieldStatus.Merged &&
                result.Value == "inner value");
            Assert.DoesNotContain("outer value", body.InnerText, System.StringComparison.Ordinal);
            Assert.Contains("inner value", body.InnerText, System.StringComparison.Ordinal);
        }

        [Fact]
        public void Test_MailMerge_RejectsOuterSimpleFieldContainingNestedComplexField() {
            using WordDocument document = WordDocument.Create();
            Body body = document._document.MainDocumentPart!.Document.Body!;
            body.Append(new Paragraph(
                new SimpleField(
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" MERGEFIELD Inner ")),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("inner placeholder")),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End })) {
                    Instruction = " MERGEFIELD Outer "
                }));

            WordTemplatePreflightReport preflight = WordMailMerge.PreflightTemplate(
                document,
                mergeFieldNames: new[] { "Outer", "Inner" });
            WordMailMergeExecutionReport execution = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["Outer"] = "outer value", ["Inner"] = "inner value" },
                removeFields: false);

            Assert.False(preflight.CanBindTemplate);
            Assert.Contains(preflight.Issues, issue =>
                issue.Kind == WordMailMergeTemplateIssueKind.MalformedMergeField &&
                issue.Name == "Outer");
            Assert.Contains(execution.Fields, result =>
                result.Instruction.Contains("Outer", System.StringComparison.Ordinal) &&
                result.Status == WordMailMergeFieldStatus.MalformedField);
            Assert.DoesNotContain("outer value", body.InnerText, System.StringComparison.Ordinal);
            Assert.Single(body.Descendants<SimpleField>());
        }

        [Fact]
        public void Test_MailMerge_UpdatesSimpleFieldResultInsideInlineWrapper() {
            using WordDocument document = WordDocument.Create();
            Body body = document._document.MainDocumentPart!.Document.Body!;
            body.Append(new Paragraph(
                new SimpleField(
                    new Hyperlink(new Run(new Text("stale result"))) { Id = "rIdMissing" }) {
                    Instruction = " MERGEFIELD Name "
                }));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["Name"] = "Ada" },
                removeFields: false);

            Assert.True(report.IsComplete);
            Assert.Equal("Ada", Assert.Single(report.Fields).Value);
            Assert.DoesNotContain("stale result", body.InnerText, System.StringComparison.Ordinal);
            Assert.Equal("Ada", body.InnerText);
            Assert.Single(body.Descendants<SimpleField>());
        }

        private static SimpleField CreateSimpleMergeFieldForFormattingTest(string name, RunProperties runProperties) {
            return new SimpleField(
                new Run(
                    (RunProperties)runProperties.CloneNode(true),
                    new Text("Placeholder"))) {
                Instruction = " MERGEFIELD  \"" + name + "\" "
            };
        }

        private static void ReplaceCellContentForFormattingTest(TableCell cell, params OpenXmlElement[] elements) {
            cell.RemoveAllChildren<Paragraph>();
            foreach (OpenXmlElement element in elements) {
                cell.Append(element);
            }
        }
    }
}
