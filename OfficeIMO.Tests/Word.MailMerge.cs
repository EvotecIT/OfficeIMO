using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_MailMerge_ReplacesFields() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMerge.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Hello ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Name\"" })
                    .AddText("!");

                var values = new Dictionary<string, string> { { "Name", "Alice" } };
                WordMailMerge.Execute(document, values);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var mainPart = document._document.MainDocumentPart;
                Assert.NotNull(mainPart);
                var xml = mainPart!.Document?.InnerText;
                Assert.NotNull(xml);
                Assert.Contains("Alice", xml);
                Assert.DoesNotContain("MERGEFIELD", xml);
            }
        }

        [Fact]
        public void Test_MailMerge_KeepFields() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeKeep.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Hello ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Name\"" })
                    .AddText("!");

                var values = new Dictionary<string, string> { { "Name", "Bob" } };
                WordMailMerge.Execute(document, values, removeFields: false);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.Fields);
                Assert.Equal("Bob", document.Fields[0].Text);
            }
        }

        [Fact]
        public void Test_MailMerge_ReplacementsPreserveFieldResultFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeFormatting.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                Body body = document._document.MainDocumentPart!.Document.Body!;
                body.Append(
                    new Paragraph(
                        new Run(new Text("Simple: ")),
                        new SimpleField(
                            new Run(
                                new RunProperties(
                                    new Bold(),
                                    new Color { Val = "C00000" }),
                                new Text("Placeholder"))) {
                            Instruction = " MERGEFIELD  \"SimpleName\" "
                        }),
                    new Paragraph(
                        new Run(new Text("Complex: ")),
                        new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                        new Run(new FieldCode(" MERGEFIELD  \"ComplexName\" ")),
                        new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                        new Run(
                            new RunProperties(
                                new Italic(),
                                new Color { Val = "008000" }),
                            new Text("Placeholder")),
                        new Run(new FieldChar { FieldCharType = FieldCharValues.End })));

                WordMailMerge.Execute(document, new Dictionary<string, string> {
                    ["SimpleName"] = "Alice",
                    ["ComplexName"] = "Bob"
                });
                document.Save(false);
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false)) {
                Body body = wordDocument.MainDocumentPart!.Document.Body!;
                Assert.Empty(body.Descendants<SimpleField>());
                Assert.Empty(body.Descendants<FieldChar>());

                Run simpleRun = Assert.Single(body.Descendants<Run>(), run => run.Elements<Text>().Any(text => text.Text == "Alice"));
                Run complexRun = Assert.Single(body.Descendants<Run>(), run => run.Elements<Text>().Any(text => text.Text == "Bob"));

                Assert.NotNull(simpleRun.RunProperties?.Bold);
                Assert.Equal("C00000", simpleRun.RunProperties!.Color!.Val!.Value);
                Assert.NotNull(complexRun.RunProperties?.Italic);
                Assert.Equal("008000", complexRun.RunProperties!.Color!.Val!.Value);
            }
        }

        [Fact]
        public void Test_MailMerge_ExecuteBatchCreatesOutputsAndKeepsTemplateUnchanged() {
            string templatePath = Path.Combine(_directoryWithFiles, "MailMergeBatchTemplate.docx");
            string outputDirectory = Path.Combine(_directoryWithFiles, "MailMergeBatch");

            using (WordDocument document = WordDocument.Create(templatePath)) {
                document.AddParagraph("Hello ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Name\"" })
                    .AddText(", your city is ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"City\"" })
                    .AddText(".");
                document.Save(false);
            }

            IReadOnlyList<string> outputs = WordMailMerge.ExecuteBatch(
                templatePath,
                new[] {
                    new Dictionary<string, string> {
                        ["Name"] = "Alice",
                        ["City"] = "Warsaw"
                    },
                    new Dictionary<string, string> {
                        ["Name"] = "Bob",
                        ["City"] = "Berlin"
                    }
                },
                (index, values) => Path.Combine(outputDirectory, $"{index + 1}-{values["Name"]}.docx"));

            Assert.Equal(2, outputs.Count);
            Assert.All(outputs, output => Assert.True(File.Exists(output)));

            using (WordDocument document = WordDocument.Load(outputs[0])) {
                string text = document._document.MainDocumentPart!.Document.InnerText;
                Assert.Contains("Alice", text);
                Assert.Contains("Warsaw", text);
                Assert.DoesNotContain("MERGEFIELD", document._document.MainDocumentPart!.Document.InnerXml);
            }

            using (WordDocument document = WordDocument.Load(outputs[1])) {
                string text = document._document.MainDocumentPart!.Document.InnerText;
                Assert.Contains("Bob", text);
                Assert.Contains("Berlin", text);
                Assert.DoesNotContain("MERGEFIELD", document._document.MainDocumentPart!.Document.InnerXml);
            }

            using (WordDocument template = WordDocument.Load(templatePath)) {
                Assert.Equal(2, template.Fields.Count(field => field.FieldType == WordFieldType.MergeField));
            }
        }

        [Fact]
        public void Test_MailMerge_RepeatsTableRowsAndRemovesTemplateRow() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeTableRows.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(2, 3);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Item";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Qty";
                table.Rows[0].Cells[2].Paragraphs[0].Text = "Price";

                table.Rows[1].Cells[0].Paragraphs[0].AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Item\"" });
                table.Rows[1].Cells[1].Paragraphs[0].AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Quantity\"" });
                table.Rows[1].Cells[2].Paragraphs[0].AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Price\"" });

                int generated = WordMailMerge.ExecuteTableRows(table, 1, new[] {
                    new Dictionary<string, string> {
                        ["Item"] = "Consulting",
                        ["Quantity"] = "2",
                        ["Price"] = "100"
                    },
                    new Dictionary<string, string> {
                        ["Item"] = "Support",
                        ["Quantity"] = "1",
                        ["Price"] = "50"
                    }
                });

                Assert.Equal(2, generated);
                Assert.Equal(3, table.Rows.Count);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordTable table = Assert.Single(document.Tables);
                Assert.Equal(3, table.Rows.Count);
                Assert.Equal("Consulting", table.Rows[1].Cells[0]._tableCell.InnerText);
                Assert.Equal("2", table.Rows[1].Cells[1]._tableCell.InnerText);
                Assert.Equal("100", table.Rows[1].Cells[2]._tableCell.InnerText);
                Assert.Equal("Support", table.Rows[2].Cells[0]._tableCell.InnerText);
                Assert.Equal("1", table.Rows[2].Cells[1]._tableCell.InnerText);
                Assert.Equal("50", table.Rows[2].Cells[2]._tableCell.InnerText);
                Assert.DoesNotContain(document.Fields, field => field.FieldType == WordFieldType.MergeField);
            }
        }

        [Fact]
        public void Test_MailMerge_RepeatsTableRowsCanKeepFields() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeTableRowsKeepFields.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].AddField(WordFieldType.MergeField, advanced: true, parameters: new List<string> { "\"Name\"" });

                WordMailMerge.ExecuteTableRows(table, 0, new[] {
                    new Dictionary<string, string> {
                        ["Name"] = "Alice"
                    },
                    new Dictionary<string, string> {
                        ["Name"] = "Bob"
                    }
                }, removeFields: false);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordTable table = Assert.Single(document.Tables);
                Assert.Equal(2, table.Rows.Count);
                Assert.EndsWith("Alice", table.Rows[0].Cells[0]._tableCell.InnerText);
                Assert.EndsWith("Bob", table.Rows[1].Cells[0]._tableCell.InnerText);
                Assert.Equal(new[] { "Alice", "Bob" }, document.Fields
                    .Where(field => field.FieldType == WordFieldType.MergeField)
                    .Select(field => field.Text)
                    .ToArray());
            }
        }

        [Fact]
        public void Test_MailMerge_RepeatsGroupedTableRowsAndPreservesFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeGroupedTableRows.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(3, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Description";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Amount";

                ReplaceCellContent(table.Rows[1].Cells[0]._tableCell,
                    new Paragraph(
                        new Run(
                            new RunProperties(new Bold(), new Color { Val = "1F4E79" }),
                            new Text("Category: ")),
                        CreateSimpleMergeField("Category", new RunProperties(new Bold(), new Color { Val = "1F4E79" }))));
                ReplaceCellContent(table.Rows[1].Cells[1]._tableCell, new Paragraph(new Run(new Text(string.Empty))));

                ReplaceCellContent(table.Rows[2].Cells[0]._tableCell,
                    new Paragraph(
                        CreateSimpleMergeField("Item", new RunProperties(new Italic(), new Color { Val = "008000" }))));
                ReplaceCellContent(table.Rows[2].Cells[1]._tableCell,
                    new Paragraph(
                        CreateSimpleMergeField("Amount", new RunProperties(new Italic(), new Color { Val = "008000" }))));

                WordMailMergeTableRowGroupResult result = WordMailMerge.ExecuteTableRowGroups(
                    table,
                    groupTemplateRowIndex: 1,
                    detailTemplateRowIndex: 2,
                    groups: new[] {
                        new WordMailMergeTableRowGroup(
                            new Dictionary<string, string> {
                                ["Category"] = "Services"
                            },
                            new[] {
                                new Dictionary<string, string> {
                                    ["Item"] = "Consulting",
                                    ["Amount"] = "100"
                                },
                                new Dictionary<string, string> {
                                    ["Item"] = "Support",
                                    ["Amount"] = "50"
                                }
                            }),
                        new WordMailMergeTableRowGroup(
                            new Dictionary<string, string> {
                                ["Category"] = "Licenses"
                            },
                            new[] {
                                new Dictionary<string, string> {
                                    ["Item"] = "Seats",
                                    ["Amount"] = "25"
                                }
                            })
                    });

                Assert.Equal(2, result.GroupCount);
                Assert.Equal(3, result.DetailRowCount);
                Assert.Equal(5, result.TotalRowCount);
                Assert.Equal(6, table.Rows.Count);
                document.Save(false);
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false)) {
                Table table = Assert.Single(wordDocument.MainDocumentPart!.Document.Body!.Descendants<Table>());
                List<TableRow> rows = table.Elements<TableRow>().ToList();
                Assert.Equal(6, rows.Count);
                Assert.Equal("DescriptionAmount", rows[0].InnerText);
                Assert.Equal("Category: Services", rows[1].InnerText);
                Assert.Equal("Consulting100", rows[2].InnerText);
                Assert.Equal("Support50", rows[3].InnerText);
                Assert.Equal("Category: Licenses", rows[4].InnerText);
                Assert.Equal("Seats25", rows[5].InnerText);
                Assert.DoesNotContain("MERGEFIELD", wordDocument.MainDocumentPart!.Document.InnerXml);

                Run groupRun = Assert.Single(rows[1].Descendants<Run>(), run => run.Elements<Text>().Any(text => text.Text == "Services"));
                Run detailRun = Assert.Single(rows[2].Descendants<Run>(), run => run.Elements<Text>().Any(text => text.Text == "Consulting"));
                Assert.NotNull(groupRun.RunProperties?.Bold);
                Assert.Equal("1F4E79", groupRun.RunProperties!.Color!.Val!.Value);
                Assert.NotNull(detailRun.RunProperties?.Italic);
                Assert.Equal("008000", detailRun.RunProperties!.Color!.Val!.Value);
            }
        }

        [Fact]
        public void Test_MailMerge_RepeatingBlocksCloneBodyContentTablesAndFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeRepeatingBlocks.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Invoice");
                document.AddParagraph("{{#each Items}}");

                Body body = document._document.MainDocumentPart!.Document.Body!;
                body.Append(new Paragraph(
                    new Run(new Text("Item: ")),
                    CreateSimpleMergeField("Item", new RunProperties(new Bold(), new Color { Val = "7030A0" }))));

                WordTable table = document.AddTable(1, 2);
                ReplaceCellContent(table.Rows[0].Cells[0]._tableCell,
                    new Paragraph(
                        new Run(new Text("Qty: ")),
                        CreateSimpleMergeField("Quantity", new RunProperties(new Italic(), new Color { Val = "008000" }))));
                ReplaceCellContent(table.Rows[0].Cells[1]._tableCell,
                    new Paragraph(
                        new Run(new Text("Price: ")),
                        CreateSimpleMergeField("Price", new RunProperties(new Italic(), new Color { Val = "008000" }))));

                document.AddParagraph("{{/each Items}}");
                document.AddParagraph("Done");

                int generated = WordMailMerge.ExecuteRepeatingBlocks(
                    document,
                    new Dictionary<string, IEnumerable<IDictionary<string, string>>> {
                        ["Items"] = new[] {
                            new Dictionary<string, string> {
                                ["Item"] = "Consulting",
                                ["Quantity"] = "2",
                                ["Price"] = "100"
                            },
                            new Dictionary<string, string> {
                                ["Item"] = "Support",
                                ["Quantity"] = "1",
                                ["Price"] = "50"
                            }
                        }
                    });

                Assert.Equal(2, generated);
                document.Save(false);
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false)) {
                Body body = wordDocument.MainDocumentPart!.Document.Body!;
                string bodyText = body.InnerText;
                Assert.Contains("Invoice", bodyText);
                Assert.Contains("Item: Consulting", bodyText);
                Assert.Contains("Qty: 2", bodyText);
                Assert.Contains("Price: 100", bodyText);
                Assert.Contains("Item: Support", bodyText);
                Assert.Contains("Qty: 1", bodyText);
                Assert.Contains("Price: 50", bodyText);
                Assert.Contains("Done", bodyText);
                Assert.DoesNotContain("{{#each Items}}", bodyText);
                Assert.DoesNotContain("{{/each Items}}", bodyText);
                Assert.DoesNotContain("MERGEFIELD", body.InnerXml);
                Assert.Equal(2, body.Descendants<Table>().Count());

                Run itemRun = Assert.Single(body.Descendants<Run>(), run => run.Elements<Text>().Any(text => text.Text == "Consulting"));
                Run quantityRun = Assert.Single(body.Descendants<Run>(), run => run.Elements<Text>().Any(text => text.Text == "2"));
                Assert.NotNull(itemRun.RunProperties?.Bold);
                Assert.Equal("7030A0", itemRun.RunProperties!.Color!.Val!.Value);
                Assert.NotNull(quantityRun.RunProperties?.Italic);
                Assert.Equal("008000", quantityRun.RunProperties!.Color!.Val!.Value);
            }
        }

        [Fact]
        public void Test_MailMerge_RepeatingBlockRegionsBindNestedData() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeNestedRepeatingBlocks.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Batch");
                document.AddParagraph("{{#each Invoices}}");

                Body body = document._document.MainDocumentPart!.Document.Body!;
                body.Append(new Paragraph(
                    new Run(new Text("Invoice: ")),
                    CreateSimpleMergeField("InvoiceNumber", new RunProperties(new Bold(), new Color { Val = "1F4E79" }))));
                document.AddParagraph("{{#each Lines}}");
                body.Append(new Paragraph(
                    new Run(new Text("Line: ")),
                    CreateSimpleMergeField("LineName", new RunProperties(new Italic(), new Color { Val = "008000" })),
                    new Run(new Text(" / Qty: ")),
                    CreateSimpleMergeField("Quantity", new RunProperties(new Italic(), new Color { Val = "008000" }))));
                document.AddParagraph("{{/each Lines}}");
                body.Append(new Paragraph(
                    new Run(new Text("Total: ")),
                    CreateSimpleMergeField("Total", new RunProperties(new Bold(), new Color { Val = "1F4E79" }))));

                document.AddParagraph("{{/each Invoices}}");
                document.AddParagraph("Done");

                int generated = WordMailMerge.ExecuteRepeatingBlockRegions(
                    document,
                    new Dictionary<string, IEnumerable<WordMailMergeBlockData>> {
                        ["Invoices"] = new[] {
                            new WordMailMergeBlockData(
                                new Dictionary<string, string> {
                                    ["InvoiceNumber"] = "INV-001",
                                    ["Total"] = "150"
                                },
                                new Dictionary<string, IEnumerable<WordMailMergeBlockData>> {
                                    ["Lines"] = new[] {
                                        new WordMailMergeBlockData(new Dictionary<string, string> {
                                            ["LineName"] = "Consulting",
                                            ["Quantity"] = "2"
                                        }),
                                        new WordMailMergeBlockData(new Dictionary<string, string> {
                                            ["LineName"] = "Support",
                                            ["Quantity"] = "1"
                                        })
                                    }
                                }),
                            new WordMailMergeBlockData(
                                new Dictionary<string, string> {
                                    ["InvoiceNumber"] = "INV-002",
                                    ["Total"] = "25"
                                },
                                new Dictionary<string, IEnumerable<WordMailMergeBlockData>> {
                                    ["Lines"] = new[] {
                                        new WordMailMergeBlockData(new Dictionary<string, string> {
                                            ["LineName"] = "Seats",
                                            ["Quantity"] = "5"
                                        })
                                    }
                                })
                        }
                    });

                Assert.Equal(5, generated);
                document.Save(false);
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false)) {
                Body body = wordDocument.MainDocumentPart!.Document.Body!;
                string bodyText = body.InnerText;
                Assert.Contains("Batch", bodyText);
                Assert.Contains("Invoice: INV-001", bodyText);
                Assert.Contains("Line: Consulting / Qty: 2", bodyText);
                Assert.Contains("Line: Support / Qty: 1", bodyText);
                Assert.Contains("Total: 150", bodyText);
                Assert.Contains("Invoice: INV-002", bodyText);
                Assert.Contains("Line: Seats / Qty: 5", bodyText);
                Assert.Contains("Total: 25", bodyText);
                Assert.Contains("Done", bodyText);
                Assert.DoesNotContain("{{#each Invoices}}", bodyText);
                Assert.DoesNotContain("{{/each Invoices}}", bodyText);
                Assert.DoesNotContain("{{#each Lines}}", bodyText);
                Assert.DoesNotContain("{{/each Lines}}", bodyText);
                Assert.DoesNotContain("MERGEFIELD", body.InnerXml);

                Run invoiceRun = Assert.Single(body.Descendants<Run>(), run => run.Elements<Text>().Any(text => text.Text == "INV-001"));
                Run lineRun = Assert.Single(body.Descendants<Run>(), run => run.Elements<Text>().Any(text => text.Text == "Consulting"));
                Assert.NotNull(invoiceRun.RunProperties?.Bold);
                Assert.Equal("1F4E79", invoiceRun.RunProperties!.Color!.Val!.Value);
                Assert.NotNull(lineRun.RunProperties?.Italic);
                Assert.Equal("008000", lineRun.RunProperties!.Color!.Val!.Value);
            }
        }

        [Fact]
        public void Test_MailMerge_ConditionalBlocksCanIncludeBodyContentAndMergeFields() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeConditionalIncluded.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Invoice");
                document.AddParagraph("{{#ShowDiscount}}");
                document.AddParagraph("Discount: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Discount\"" });
                WordTable table = document.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Discount note";
                document.AddParagraph("{{/ShowDiscount}}");
                document.AddParagraph("Footer");

                int processed = WordMailMerge.ExecuteConditionalBlocks(document, new Dictionary<string, bool> {
                    ["ShowDiscount"] = true
                });
                WordMailMerge.Execute(document, new Dictionary<string, string> {
                    ["Discount"] = "10%"
                });

                Assert.Equal(1, processed);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                string bodyText = document._document.MainDocumentPart!.Document.Body!.InnerText;
                Assert.Contains("Invoice", bodyText);
                Assert.Contains("Discount: 10%", bodyText);
                Assert.Contains("Discount note", bodyText);
                Assert.Contains("Footer", bodyText);
                Assert.DoesNotContain("{{#ShowDiscount}}", bodyText);
                Assert.DoesNotContain("{{/ShowDiscount}}", bodyText);
                Assert.Single(document.Tables);
                Assert.DoesNotContain(document.Fields, field => field.FieldType == WordFieldType.MergeField);
            }
        }

        [Fact]
        public void Test_MailMerge_ConditionalBlocksCanRemoveBodyContent() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeConditionalRemoved.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Invoice");
                document.AddParagraph("{{#ShowDiscount}}");
                document.AddParagraph("Discount: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Discount\"" });
                WordTable table = document.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Discount note";
                document.AddParagraph("{{/ShowDiscount}}");
                document.AddParagraph("Footer");

                int processed = WordMailMerge.ExecuteConditionalBlocks(document, new Dictionary<string, bool> {
                    ["ShowDiscount"] = false
                });

                Assert.Equal(1, processed);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                string bodyText = document._document.MainDocumentPart!.Document.Body!.InnerText;
                Assert.Contains("Invoice", bodyText);
                Assert.Contains("Footer", bodyText);
                Assert.DoesNotContain("Discount", bodyText);
                Assert.DoesNotContain("{{#ShowDiscount}}", bodyText);
                Assert.DoesNotContain("{{/ShowDiscount}}", bodyText);
                Assert.Empty(document.Tables);
                Assert.DoesNotContain(document.Fields, field => field.FieldType == WordFieldType.MergeField);
            }
        }

        [Fact]
        public void Test_MailMerge_ConditionalBlocksThrowForMissingCondition() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeConditionalMissing.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("{{#ShowDiscount}}");
                document.AddParagraph("Discount");
                document.AddParagraph("{{/ShowDiscount}}");

                var exception = Assert.Throws<InvalidOperationException>(() =>
                    WordMailMerge.ExecuteConditionalBlocks(document, new Dictionary<string, bool>()));
                Assert.Contains("ShowDiscount", exception.Message);
            }
        }

        [Fact]
        public void Test_MailMerge_ConditionalBlocksHandleNestedRegions() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeConditionalNested.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("{{#ShowSummary}}");
                document.AddParagraph("Summary");
                document.AddParagraph("{{#ShowDetails}}");
                document.AddParagraph("Details");
                document.AddParagraph("{{/ShowDetails}}");
                document.AddParagraph("Total");
                document.AddParagraph("{{/ShowSummary}}");

                int processed = WordMailMerge.ExecuteConditionalBlocks(document, new Dictionary<string, bool> {
                    ["ShowSummary"] = true,
                    ["ShowDetails"] = false
                });

                Assert.Equal(2, processed);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                string bodyText = document._document.MainDocumentPart!.Document.Body!.InnerText;
                Assert.Contains("Summary", bodyText);
                Assert.Contains("Total", bodyText);
                Assert.DoesNotContain("Details", bodyText);
                Assert.DoesNotContain("{{#ShowSummary}}", bodyText);
                Assert.DoesNotContain("{{/ShowSummary}}", bodyText);
                Assert.DoesNotContain("{{#ShowDetails}}", bodyText);
                Assert.DoesNotContain("{{/ShowDetails}}", bodyText);
            }
        }

        [Fact]
        public void Test_MailMerge_ConditionalBlocksCanIncludeTableCellContentAndMergeFields() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeConditionalTableCellIncluded.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "{{#ShowDetail}}";
                table.Rows[0].Cells[0].AddParagraph("Detail: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Detail\"" });
                table.Rows[0].Cells[0].AddParagraph("{{/ShowDetail}}");
                table.Rows[0].Cells[0].AddParagraph("Always visible");

                int processed = WordMailMerge.ExecuteConditionalBlocks(document, new Dictionary<string, bool> {
                    ["ShowDetail"] = true
                });
                WordMailMerge.Execute(document, new Dictionary<string, string> {
                    ["Detail"] = "Approved"
                });

                Assert.Equal(1, processed);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                string bodyText = document._document.MainDocumentPart!.Document.Body!.InnerText;
                Assert.Contains("Detail: Approved", bodyText);
                Assert.Contains("Always visible", bodyText);
                Assert.DoesNotContain("{{#ShowDetail}}", bodyText);
                Assert.DoesNotContain("{{/ShowDetail}}", bodyText);
                Assert.Single(document.Tables);
                Assert.DoesNotContain(document.Fields, field => field.FieldType == WordFieldType.MergeField);
            }
        }

        [Fact]
        public void Test_MailMerge_ConditionalBlocksCanRemoveTableCellContent() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeConditionalTableCellRemoved.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "{{#ShowDetail}}";
                table.Rows[0].Cells[0].AddParagraph("Detail")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Detail\"" });
                table.Rows[0].Cells[0].AddParagraph("{{/ShowDetail}}");
                table.Rows[0].Cells[0].AddParagraph("Always visible");

                int processed = WordMailMerge.ExecuteConditionalBlocks(document, new Dictionary<string, bool> {
                    ["ShowDetail"] = false
                });

                Assert.Equal(1, processed);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                string bodyText = document._document.MainDocumentPart!.Document.Body!.InnerText;
                Assert.DoesNotContain("Detail", bodyText);
                Assert.Contains("Always visible", bodyText);
                Assert.DoesNotContain("{{#ShowDetail}}", bodyText);
                Assert.DoesNotContain("{{/ShowDetail}}", bodyText);
                Assert.Single(document.Tables);
                Assert.DoesNotContain(document.Fields, field => field.FieldType == WordFieldType.MergeField);
            }
        }

        [Fact]
        public void Test_MailMerge_InspectTemplateReportsFieldsAndConditionalBlocks() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeTemplateInspect.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Hello ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Name\"" });
                document.AddParagraph("{{#ShowDiscount}}");
                document.AddParagraph("Discount: ")
                    .AddField(WordFieldType.MergeField, advanced: true, parameters: new List<string> { "\"Discount\"" });
                document.AddParagraph("{{/ShowDiscount}}");

                var inspection = WordMailMerge.InspectTemplate(
                    document,
                    mergeFieldNames: new[] { "Name", "Discount" },
                    conditionNames: new[] { "ShowDiscount" });

                Assert.True(inspection.IsValid);
                Assert.Equal(new[] { "Discount", "Name" }, inspection.MergeFieldNames);
                Assert.Equal(new[] { "ShowDiscount" }, inspection.ConditionalBlockNames);
                Assert.Same(inspection, inspection.EnsureValid());
            }
        }

        [Fact]
        public void Test_MailMerge_InspectTemplateReportsMissingBindingsAndBrokenBlocks() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeTemplateInspectMissing.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Hello ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Name\"" });
                document.AddParagraph("{{#ShowDiscount}}");
                document.AddParagraph("Discount");

                var inspection = WordMailMerge.InspectTemplate(
                    document,
                    mergeFieldNames: new[] { "Other" },
                    conditionNames: new[] { "OtherCondition" });

                Assert.False(inspection.IsValid);
                Assert.Contains("Name", inspection.MergeFieldNames);
                Assert.Contains("ShowDiscount", inspection.ConditionalBlockNames);
                Assert.Contains(inspection.Issues, issue => issue.Kind == WordMailMergeTemplateIssueKind.MissingMergeFieldValue && issue.Name == "Name");
                Assert.Contains(inspection.Issues, issue => issue.Kind == WordMailMergeTemplateIssueKind.MissingConditionalValue && issue.Name == "ShowDiscount");
                Assert.Contains(inspection.Issues, issue => issue.Kind == WordMailMergeTemplateIssueKind.UnmatchedConditionalStart && issue.Name == "ShowDiscount");

                var exception = Assert.Throws<InvalidOperationException>(() => inspection.EnsureValid());
                Assert.Contains("ShowDiscount", exception.Message);
            }
        }

        [Fact]
        public void Test_MailMerge_InspectTemplateReportsConditionalBlocksInsideTableCells() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeTemplateInspectTableCell.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "{{#ShowDetail}}";
                table.Rows[0].Cells[0].AddParagraph("Detail");
                table.Rows[0].Cells[0].AddParagraph("{{/ShowDetail}}");

                var inspection = WordMailMerge.InspectTemplate(
                    document,
                    mergeFieldNames: new string[0],
                    conditionNames: new[] { "ShowDetail" });

                Assert.True(inspection.IsValid);
                Assert.Equal(new[] { "ShowDetail" }, inspection.ConditionalBlockNames);
            }
        }

        [Fact]
        public void Test_MailMerge_InspectTemplateReportsRepeatingBlocks() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeTemplateInspectRepeatingBlocks.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("{{#each Invoices}}");
                document.AddParagraph("Invoice ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"InvoiceNumber\"" });
                document.AddParagraph("{{#each Lines}}");
                document.AddParagraph("Line ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"LineName\"" });
                document.AddParagraph("{{/each Lines}}");
                document.AddParagraph("{{/each Invoices}}");

                var inspection = WordMailMerge.InspectTemplate(
                    document,
                    mergeFieldNames: new[] { "InvoiceNumber", "LineName" },
                    conditionNames: new string[0],
                    repeatingBlockNames: new[] { "Invoices", "Lines" });

                Assert.True(inspection.IsValid);
                Assert.Equal(new[] { "InvoiceNumber", "LineName" }, inspection.MergeFieldNames);
                Assert.Empty(inspection.ConditionalBlockNames);
                Assert.Equal(new[] { "Invoices", "Lines" }, inspection.RepeatingBlockNames);
            }
        }

        [Fact]
        public void Test_MailMerge_InspectTemplateReportsMissingAndBrokenRepeatingBlocks() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeTemplateInspectRepeatingBlocksMissing.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("{{#each Invoices}}");
                document.AddParagraph("Invoice");
                document.AddParagraph("{{#each Lines}}");
                document.AddParagraph("Line");
                document.AddParagraph("{{/each OtherLines}}");

                var inspection = WordMailMerge.InspectTemplate(
                    document,
                    mergeFieldNames: new string[0],
                    conditionNames: new string[0],
                    repeatingBlockNames: new[] { "Invoices" });

                Assert.False(inspection.IsValid);
                Assert.Contains("Invoices", inspection.RepeatingBlockNames);
                Assert.Contains("Lines", inspection.RepeatingBlockNames);
                Assert.Contains("OtherLines", inspection.RepeatingBlockNames);
                Assert.Contains(inspection.Issues, issue => issue.Kind == WordMailMergeTemplateIssueKind.MissingRepeatingBlockData && issue.Name == "Lines");
                Assert.Contains(inspection.Issues, issue => issue.Kind == WordMailMergeTemplateIssueKind.MissingRepeatingBlockData && issue.Name == "OtherLines");
                Assert.Contains(inspection.Issues, issue => issue.Kind == WordMailMergeTemplateIssueKind.MismatchedRepeatingBlockEnd && issue.Name == "OtherLines");
                Assert.Contains(inspection.Issues, issue => issue.Kind == WordMailMergeTemplateIssueKind.UnmatchedRepeatingBlockStart && issue.Name == "Invoices");

                var exception = Assert.Throws<InvalidOperationException>(() => inspection.EnsureValid());
                Assert.Contains("Repeating block", exception.Message);
            }
        }

        [Fact]
        public void Test_MailMerge_ConditionalBlocksCanRunInsideBlockContentControls() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeConditionalContentControl.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var block = new SdtBlock(
                    new SdtProperties(
                        new SdtAlias { Val = "Panel" },
                        new Tag { Val = "Panel" },
                        new SdtId { Val = 42 }),
                    new SdtContentBlock(
                        CreateMailMergeParagraph("{{#ShowPanel}}"),
                        CreateMailMergeParagraph("Panel visible"),
                        CreateMailMergeParagraph("{{/ShowPanel}}"),
                        CreateMailMergeParagraph("Always visible")));

                document._document.MainDocumentPart!.Document.Body!.Append(block);

                var inspection = WordMailMerge.InspectTemplate(
                    document,
                    mergeFieldNames: new string[0],
                    conditionNames: new[] { "ShowPanel" });

                Assert.True(inspection.IsValid);
                Assert.Equal(new[] { "ShowPanel" }, inspection.ConditionalBlockNames);

                int processed = WordMailMerge.ExecuteConditionalBlocks(document, new Dictionary<string, bool> {
                    ["ShowPanel"] = false
                });

                Assert.Equal(1, processed);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                string bodyText = document._document.MainDocumentPart!.Document.Body!.InnerText;
                Assert.DoesNotContain("Panel visible", bodyText);
                Assert.Contains("Always visible", bodyText);
                Assert.DoesNotContain("{{#ShowPanel}}", bodyText);
                Assert.DoesNotContain("{{/ShowPanel}}", bodyText);
                Assert.Single(document.StructuredDocumentTags);
            }
        }

        [Fact]
        public void Test_MailMerge_ConditionalBlocksCanRunInsideHeadersAndFooters() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeConditionalHeaderFooter.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Body");
                document.AddHeadersAndFooters();

                var header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                header.AddParagraph("{{#ShowHeader}}");
                header.AddParagraph("Header secret");
                header.AddParagraph("{{/ShowHeader}}");
                header.AddParagraph("Header stable");

                var footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
                footer.AddParagraph("{{#ShowFooter}}");
                footer.AddParagraph("Footer visible");
                footer.AddParagraph("{{/ShowFooter}}");

                var inspection = WordMailMerge.InspectTemplate(
                    document,
                    mergeFieldNames: new string[0],
                    conditionNames: new[] { "ShowHeader", "ShowFooter" });

                Assert.True(inspection.IsValid);
                Assert.Equal(new[] { "ShowFooter", "ShowHeader" }, inspection.ConditionalBlockNames);

                int processed = WordMailMerge.ExecuteConditionalBlocks(document, new Dictionary<string, bool> {
                    ["ShowHeader"] = false,
                    ["ShowFooter"] = true
                });

                Assert.Equal(2, processed);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                var footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
                string headerText = string.Concat(header.Paragraphs.Select(paragraph => paragraph.Text));
                string footerText = string.Concat(footer.Paragraphs.Select(paragraph => paragraph.Text));

                Assert.DoesNotContain("Header secret", headerText);
                Assert.Contains("Header stable", headerText);
                Assert.Contains("Footer visible", footerText);
                Assert.DoesNotContain("{{#ShowHeader}}", headerText);
                Assert.DoesNotContain("{{/ShowHeader}}", headerText);
                Assert.DoesNotContain("{{#ShowFooter}}", footerText);
                Assert.DoesNotContain("{{/ShowFooter}}", footerText);
            }
        }

        [Fact]
        public void Test_MailMerge_RefreshesContentControlDataBindingsFromCustomXml() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeContentControlBindingRefresh.docx");
            const string storeItemId = "{11111111-2222-3333-4444-555555555555}";
            const string schemaUri = "urn:officeimo:test:client";

            using (WordDocument document = WordDocument.Create(filePath)) {
                AddClientCustomXmlPart(document, storeItemId, schemaUri, "Alice");
                document._document.MainDocumentPart!.Document.Body!.Append(CreateBoundClientContentControl(storeItemId, schemaUri, "Placeholder"));

                WordContentControlDataBindingResult result = WordMailMerge.RefreshContentControlDataBindings(document);

                Assert.Equal(1, result.BindingCount);
                Assert.Equal(1, result.UpdatedContentControls);
                Assert.Equal(0, result.UpdatedCustomXmlNodes);
                Assert.False(result.HasMissingValues);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.StructuredDocumentTags);
                Assert.Equal("Alice", document.StructuredDocumentTags[0].Text);
            }
        }

        [Fact]
        public void Test_MailMerge_ContentControlBindingDoesNotFallbackWhenStoreItemIdIsMissing() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeContentControlMissingStoreItem.docx");
            const string existingStoreItemId = "{33333333-4444-5555-6666-777777777777}";
            const string missingStoreItemId = "{44444444-5555-6666-7777-888888888888}";
            const string schemaUri = "urn:officeimo:test:client";

            using (WordDocument document = WordDocument.Create(filePath)) {
                AddClientCustomXmlPart(document, existingStoreItemId, schemaUri, "Alice");
                document._document.MainDocumentPart!.Document.Body!.Append(CreateBoundClientContentControl(missingStoreItemId, schemaUri, "Placeholder"));

                WordContentControlDataBindingResult refreshResult = WordMailMerge.RefreshContentControlDataBindings(document);

                Assert.Equal(1, refreshResult.BindingCount);
                Assert.Equal(0, refreshResult.UpdatedContentControls);
                Assert.Equal(0, refreshResult.UpdatedCustomXmlNodes);
                Assert.True(refreshResult.HasMissingValues);
                Assert.Contains("ClientName", refreshResult.MissingValueKeys);

                WordContentControlDataBindingResult executeResult = WordMailMerge.ExecuteContentControlDataBindings(
                    document,
                    new Dictionary<string, string> {
                        ["ClientName"] = "Bob"
                    });

                Assert.Equal(1, executeResult.BindingCount);
                Assert.Equal(1, executeResult.UpdatedContentControls);
                Assert.Equal(0, executeResult.UpdatedCustomXmlNodes);
                Assert.False(executeResult.HasMissingValues);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("Bob", document.StructuredDocumentTags[0].Text);

                CustomXmlPart customXmlPart = Assert.Single(document._wordprocessingDocument.MainDocumentPart!.CustomXmlParts);
                using Stream stream = customXmlPart.GetStream(FileMode.Open, FileAccess.Read);
                XDocument xml = XDocument.Load(stream);
                XNamespace ns = XNamespace.Get(schemaUri);
                Assert.Equal("Alice", xml.Root?.Element(ns + "Client")?.Element(ns + "Name")?.Value);
            }
        }

        [Fact]
        public void Test_MailMerge_ContentControlBindingSupportsCustomXmlAttributes() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeContentControlAttributeBinding.docx");
            const string storeItemId = "{55555555-6666-7777-8888-999999999999}";
            const string schemaUri = "urn:officeimo:test:client";

            using (WordDocument document = WordDocument.Create(filePath)) {
                AddClientAttributeCustomXmlPart(document, storeItemId, schemaUri, "Alice");
                document._document.MainDocumentPart!.Document.Body!.Append(CreateBoundClientAttributeContentControl(storeItemId, schemaUri, "Placeholder"));

                WordContentControlDataBindingResult refreshResult = WordMailMerge.RefreshContentControlDataBindings(document);

                Assert.Equal(1, refreshResult.BindingCount);
                Assert.Equal(1, refreshResult.UpdatedContentControls);
                Assert.False(refreshResult.HasMissingValues);

                WordContentControlDataBindingResult executeResult = WordMailMerge.ExecuteContentControlDataBindings(
                    document,
                    new Dictionary<string, string> {
                        ["ClientAttributeName"] = "Bob"
                    });

                Assert.Equal(1, executeResult.UpdatedContentControls);
                Assert.Equal(1, executeResult.UpdatedCustomXmlNodes);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("Bob", document.StructuredDocumentTags[0].Text);

                CustomXmlPart customXmlPart = Assert.Single(document._wordprocessingDocument.MainDocumentPart!.CustomXmlParts);
                using Stream stream = customXmlPart.GetStream(FileMode.Open, FileAccess.Read);
                XDocument xml = XDocument.Load(stream);
                XNamespace ns = XNamespace.Get(schemaUri);
                Assert.Equal("Bob", xml.Root?.Element(ns + "Client")?.Attribute("name")?.Value);
            }
        }

        [Fact]
        public void Test_MailMerge_ContentControlBindingRefreshesRowControls() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeContentControlRowBinding.docx");
            const string storeItemId = "{66666666-7777-8888-9999-000000000000}";
            const string schemaUri = "urn:officeimo:test:client";

            using (WordDocument document = WordDocument.Create(filePath)) {
                AddClientCustomXmlPart(document, storeItemId, schemaUri, "Alice");
                document._document.MainDocumentPart!.Document.Body!.Append(new Table(CreateBoundClientRowContentControl(storeItemId, schemaUri, "Placeholder")));

                WordContentControlDataBindingResult result = WordMailMerge.RefreshContentControlDataBindings(document);

                Assert.Equal(1, result.BindingCount);
                Assert.Equal(1, result.UpdatedContentControls);
                Assert.False(result.HasMissingValues);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                SdtRow row = Assert.Single(document._document.MainDocumentPart!.Document.Body!.Descendants<SdtRow>());
                Assert.Equal("Alice", string.Concat(row.Descendants<Text>().Select(text => text.Text)));
            }
        }

        [Fact]
        public void Test_MailMerge_ContentControlBindingRefreshesRowControlWithoutClearingSiblingCells() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeContentControlRowBindingSiblingCells.docx");
            const string storeItemId = "{77777777-8888-9999-0000-111111111111}";
            const string schemaUri = "urn:officeimo:test:client";

            using (WordDocument document = WordDocument.Create(filePath)) {
                AddClientCustomXmlPart(document, storeItemId, schemaUri, "Alice");
                SdtRow control = CreateBoundClientRowContentControl(storeItemId, schemaUri, "Placeholder", "Keep me");
                control.Descendants<TableCell>().First().Descendants<Paragraph>().First()
                    .Append(new Run(new Text(" stale") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));
                document._document.MainDocumentPart!.Document.Body!.Append(new Table(control));

                WordContentControlDataBindingResult result = WordMailMerge.RefreshContentControlDataBindings(document);

                Assert.Equal(1, result.BindingCount);
                Assert.Equal(1, result.UpdatedContentControls);
                Assert.False(result.HasMissingValues);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                SdtRow row = Assert.Single(document._document.MainDocumentPart!.Document.Body!.Descendants<SdtRow>());
                TableCell[] cells = row.Descendants<TableCell>().ToArray();
                Assert.Equal("Alice", cells[0].InnerText);
                Assert.Equal("Keep me", cells[1].InnerText);
            }
        }

        [Fact]
        public void Test_MailMerge_ExecutesContentControlDataBindingsAndUpdatesCustomXml() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeContentControlBindingExecute.docx");
            const string storeItemId = "{22222222-3333-4444-5555-666666666666}";
            const string schemaUri = "urn:officeimo:test:client";

            using (WordDocument document = WordDocument.Create(filePath)) {
                AddClientCustomXmlPart(document, storeItemId, schemaUri, "Alice");
                document._document.MainDocumentPart!.Document.Body!.Append(CreateBoundClientContentControl(storeItemId, schemaUri, "Placeholder"));

                WordContentControlDataBindingResult result = WordMailMerge.ExecuteContentControlDataBindings(
                    document,
                    new Dictionary<string, string> {
                        ["ClientName"] = "Bob"
                    });

                Assert.Equal(1, result.BindingCount);
                Assert.Equal(1, result.UpdatedContentControls);
                Assert.Equal(1, result.UpdatedCustomXmlNodes);
                Assert.False(result.HasMissingValues);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("Bob", document.StructuredDocumentTags[0].Text);

                CustomXmlPart customXmlPart = Assert.Single(document._wordprocessingDocument.MainDocumentPart!.CustomXmlParts);
                using Stream stream = customXmlPart.GetStream(FileMode.Open, FileAccess.Read);
                XDocument xml = XDocument.Load(stream);
                XNamespace ns = XNamespace.Get(schemaUri);
                Assert.Equal("Bob", xml.Root?.Element(ns + "Client")?.Element(ns + "Name")?.Value);
            }
        }

        private static Paragraph CreateMailMergeParagraph(string text) {
            return new Paragraph(new Run(new Text(text)));
        }

        private static SimpleField CreateSimpleMergeField(string name, RunProperties runProperties) {
            return new SimpleField(
                new Run(
                    (RunProperties)runProperties.CloneNode(true),
                    new Text("Placeholder"))) {
                Instruction = $" MERGEFIELD \"{name}\" "
            };
        }

        private static void ReplaceCellContent(TableCell cell, Paragraph paragraph) {
            cell.RemoveAllChildren<Paragraph>();
            cell.Append(paragraph);
        }

        private static void AddClientCustomXmlPart(WordDocument document, string storeItemId, string schemaUri, string name) {
            MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;
            CustomXmlPart customXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);

            using (Stream stream = customXmlPart.GetStream(FileMode.Create, FileAccess.Write)) {
                var xml = new XDocument(
                    new XElement(XName.Get("Root", schemaUri),
                        new XElement(XName.Get("Client", schemaUri),
                            new XElement(XName.Get("Name", schemaUri), name))));
                xml.Save(stream);
            }

            CustomXmlPropertiesPart propertiesPart = customXmlPart.AddNewPart<CustomXmlPropertiesPart>();
            propertiesPart.DataStoreItem = new DataStoreItem { ItemId = storeItemId };
            propertiesPart.DataStoreItem.Append(new SchemaReferences(new SchemaReference { Uri = schemaUri }));
            propertiesPart.DataStoreItem.Save();
        }

        private static void AddClientAttributeCustomXmlPart(WordDocument document, string storeItemId, string schemaUri, string name) {
            MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;
            CustomXmlPart customXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);

            using (Stream stream = customXmlPart.GetStream(FileMode.Create, FileAccess.Write)) {
                var xml = new XDocument(
                    new XElement(XName.Get("Root", schemaUri),
                        new XElement(XName.Get("Client", schemaUri),
                            new XAttribute("name", name))));
                xml.Save(stream);
            }

            CustomXmlPropertiesPart propertiesPart = customXmlPart.AddNewPart<CustomXmlPropertiesPart>();
            propertiesPart.DataStoreItem = new DataStoreItem { ItemId = storeItemId };
            propertiesPart.DataStoreItem.Append(new SchemaReferences(new SchemaReference { Uri = schemaUri }));
            propertiesPart.DataStoreItem.Save();
        }

        private static SdtBlock CreateBoundClientContentControl(string storeItemId, string schemaUri, string text) {
            return new SdtBlock(
                new SdtProperties(
                    new SdtAlias { Val = "ClientName" },
                    new Tag { Val = "Client.Name" },
                    new SdtId { Val = 1001 },
                    new DataBinding {
                        PrefixMappings = $"xmlns:c='{schemaUri}'",
                        XPath = "/c:Root[1]/c:Client[1]/c:Name[1]",
                        StoreItemId = storeItemId
                    },
                    new SdtContentText()),
                new SdtContentBlock(CreateMailMergeParagraph(text)));
        }

        private static SdtBlock CreateBoundClientAttributeContentControl(string storeItemId, string schemaUri, string text) {
            return new SdtBlock(
                new SdtProperties(
                    new SdtAlias { Val = "ClientAttributeName" },
                    new Tag { Val = "Client.AttributeName" },
                    new SdtId { Val = 1002 },
                    new DataBinding {
                        PrefixMappings = $"xmlns:c='{schemaUri}'",
                        XPath = "/c:Root[1]/c:Client[1]/@name",
                        StoreItemId = storeItemId
                    },
                    new SdtContentText()),
                new SdtContentBlock(CreateMailMergeParagraph(text)));
        }

        private static SdtRow CreateBoundClientRowContentControl(string storeItemId, string schemaUri, string text, string? siblingCellText = null) {
            var row = new TableRow(new TableCell(CreateMailMergeParagraph(text)));
            if (siblingCellText != null) {
                row.Append(new TableCell(CreateMailMergeParagraph(siblingCellText)));
            }

            return new SdtRow(
                new SdtProperties(
                    new SdtAlias { Val = "ClientName" },
                    new Tag { Val = "Client.Name.Row" },
                    new SdtId { Val = 1003 },
                    new DataBinding {
                        PrefixMappings = $"xmlns:c='{schemaUri}'",
                        XPath = "/c:Root[1]/c:Client[1]/c:Name[1]",
                        StoreItemId = storeItemId
                    },
                    new SdtContentText()),
                new SdtContentRow(row));
        }
    }
}
