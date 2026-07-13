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
