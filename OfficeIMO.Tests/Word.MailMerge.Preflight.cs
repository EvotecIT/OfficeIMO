using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_MailMerge_PreflightTemplateReportsCapabilitiesAndSerializes() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergePreflightClean.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Hello ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Name\"" });
                document.AddParagraph("{{#ShowDiscount}}");
                document.AddParagraph("Discount: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Discount\"" });
                document.AddParagraph("{{/ShowDiscount}}");
                document.AddParagraph("{{#each Lines}}");
                document.AddParagraph("Line ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"LineName\"" });
                document.AddParagraph("{{/each Lines}}");

                WordTemplatePreflightReport report = WordMailMerge.PreflightTemplate(
                    document,
                    mergeFieldNames: new[] { "Name", "Discount", "LineName" },
                    conditionNames: new[] { "ShowDiscount" },
                    repeatingBlockNames: new[] { "Lines" });

                Assert.True(report.CanBindTemplate);
                Assert.True(report.Can(WordTemplatePreflightCapability.BindMergeFields));
                Assert.True(report.Can(WordTemplatePreflightCapability.BindConditionalBlocks));
                Assert.True(report.Can(WordTemplatePreflightCapability.BindRepeatingBlocks));
                Assert.Same(report, report.EnsureCan(WordTemplatePreflightCapability.BindTemplate));
                Assert.Equal(3, report.MergeFieldCount);
                Assert.Equal(1, report.ConditionalBlockCount);
                Assert.Equal(1, report.RepeatingBlockCount);
                Assert.Equal(0, report.IssueCount);

                string json = report.ToJson();
                Assert.Contains("\"canBindTemplate\": true", json);
                Assert.Contains("\"mergeFieldNames\": [\"Discount\", \"LineName\", \"Name\"]", json);

                string markdown = report.ToMarkdown();
                Assert.Contains("# Word Template Preflight Report", markdown);
                Assert.Contains("| Can bind template | yes |", markdown);
            }
        }

        [Fact]
        public void Test_MailMerge_PreflightTemplateSeparatesCapabilityDiagnostics() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergePreflightBlocked.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Hello ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Name\"" });
                document.AddParagraph("{{#ShowDiscount}}");
                document.AddParagraph("{{#each Lines}}");
                document.AddParagraph("Line");

                WordTemplatePreflightReport report = WordMailMerge.PreflightTemplate(
                    document,
                    mergeFieldNames: new string[0],
                    conditionNames: new string[0],
                    repeatingBlockNames: new string[0]);

                Assert.False(report.CanBindTemplate);
                Assert.False(report.Can(WordTemplatePreflightCapability.BindMergeFields));
                Assert.False(report.Can(WordTemplatePreflightCapability.BindConditionalBlocks));
                Assert.False(report.Can(WordTemplatePreflightCapability.BindRepeatingBlocks));

                Assert.Contains(report.GetDiagnostics(WordTemplatePreflightCapability.BindMergeFields), issue =>
                    issue.Kind == WordMailMergeTemplateIssueKind.MissingMergeFieldValue && issue.Name == "Name");
                Assert.Contains(report.GetDiagnostics(WordTemplatePreflightCapability.BindConditionalBlocks), issue =>
                    issue.Kind == WordMailMergeTemplateIssueKind.UnmatchedConditionalStart && issue.Name == "ShowDiscount");
                Assert.Contains(report.GetDiagnostics(WordTemplatePreflightCapability.BindRepeatingBlocks), issue =>
                    issue.Kind == WordMailMergeTemplateIssueKind.UnmatchedRepeatingBlockStart && issue.Name == "Lines");

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                    report.EnsureCan(WordTemplatePreflightCapability.BindMergeFields));
                Assert.Contains("Merge field 'Name' was not supplied.", exception.Message);

                string markdown = report.ToMarkdown();
                Assert.Contains("MissingMergeFieldValue", markdown);
                Assert.Contains("UnmatchedConditionalStart", markdown);
                Assert.Contains("UnmatchedRepeatingBlockStart", markdown);
            }
        }

        [Fact]
        public void Test_MailMerge_PreflightTemplateSeesTableCellTemplateMarkersAfterSaveLoad() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergePreflightTableCellMarkers.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(1, 2);
                ReplaceTableCellPreflightContent(
                    table.Rows[0].Cells[0]._tableCell,
                    new Paragraph(new Run(new Text("{{#ShowLine}}"))),
                    new Paragraph(new Run(new Text("Line ")), CreatePreflightMergeField("LineName")),
                    new Paragraph(new Run(new Text("{{/ShowLine}}"))));
                ReplaceTableCellPreflightContent(
                    table.Rows[0].Cells[1]._tableCell,
                    new Paragraph(new Run(new Text("{{#each Tasks}}"))),
                    new Paragraph(new Run(new Text("Task ")), CreatePreflightMergeField("TaskName")),
                    new Paragraph(new Run(new Text("{{/each Tasks}}"))));
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());

                WordTemplatePreflightReport report = WordMailMerge.PreflightTemplate(
                    document,
                    mergeFieldNames: new[] { "LineName", "TaskName" },
                    conditionNames: new[] { "ShowLine" },
                    repeatingBlockNames: new[] { "Tasks" });

                Assert.True(report.CanBindTemplate);
                Assert.Equal(2, report.MergeFieldCount);
                Assert.Equal(1, report.ConditionalBlockCount);
                Assert.Equal(1, report.RepeatingBlockCount);
                Assert.Equal(0, report.IssueCount);

                string json = report.ToJson();
                Assert.Contains("\"conditionalBlockNames\": [\"ShowLine\"]", json);
                Assert.Contains("\"repeatingBlockNames\": [\"Tasks\"]", json);

                string markdown = report.ToMarkdown();
                Assert.Contains("| Merge fields | 2 |", markdown);
                Assert.Contains("| Conditional blocks | 1 |", markdown);
                Assert.Contains("| Repeating blocks | 1 |", markdown);
            }
        }

        [Fact]
        public void Test_MailMerge_PreflightTemplateSeesHeaderFooterTemplateMarkersAfterSaveLoad() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergePreflightHeaderFooterMarkers.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Body");
                document.AddHeadersAndFooters();

                WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                header.AddParagraph("{{#ShowClientHeader}}");
                header.AddParagraph("Client: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"ClientName\"" });
                header.AddParagraph("{{/ShowClientHeader}}");

                WordFooter footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
                footer.AddParagraph("{{#each Signers}}");
                footer.AddParagraph("Signer: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"SignerName\"" });
                footer.AddParagraph("{{/each Signers}}");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());

                WordTemplatePreflightReport report = WordMailMerge.PreflightTemplate(
                    document,
                    mergeFieldNames: new[] { "ClientName", "SignerName" },
                    conditionNames: new[] { "ShowClientHeader" },
                    repeatingBlockNames: new[] { "Signers" });

                Assert.True(report.CanBindTemplate);
                Assert.Equal(2, report.MergeFieldCount);
                Assert.Equal(1, report.ConditionalBlockCount);
                Assert.Equal(1, report.RepeatingBlockCount);
                Assert.Equal(0, report.IssueCount);

                string json = report.ToJson();
                Assert.Contains("\"mergeFieldNames\": [\"ClientName\", \"SignerName\"]", json);
                Assert.Contains("\"conditionalBlockNames\": [\"ShowClientHeader\"]", json);
                Assert.Contains("\"repeatingBlockNames\": [\"Signers\"]", json);

                string markdown = report.ToMarkdown();
                Assert.Contains("| Merge fields | 2 |", markdown);
                Assert.Contains("| Conditional blocks | 1 |", markdown);
                Assert.Contains("| Repeating blocks | 1 |", markdown);
            }
        }

        [Fact]
        public void Test_MailMerge_PreflightTemplateReportsUnsupportedWordNativeRecordControlFields() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergePreflightWordRecordControlFields.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Name: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Name\"" });
                document.AddParagraph("Next active: ")._paragraph.Append(new SimpleField(new Run(new Text("stale-next"))) {
                    Instruction = " NEXTIF \"Status\" = \"Active\" "
                });
                document._document.Body!.Append(CreatePreflightComplexFieldParagraph(" SKIPIF \"Region\" <> \"EU\" ", "stale-skip"));
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordTemplatePreflightReport report = WordMailMerge.PreflightTemplate(
                    document,
                    mergeFieldNames: new[] { "Name" });

                Assert.False(report.CanBindTemplate);
                Assert.True(report.Can(WordTemplatePreflightCapability.BindMergeFields));
                Assert.Equal(1, report.MergeFieldCount);
                Assert.Equal(2, report.IssueCount);
                Assert.Contains(report.Issues, issue =>
                    issue.Kind == WordMailMergeTemplateIssueKind.UnsupportedMailMergeControlField &&
                    issue.Name == "NEXTIF" &&
                    issue.Message.Contains("NEXTIF", StringComparison.Ordinal));
                Assert.Contains(report.Issues, issue =>
                    issue.Kind == WordMailMergeTemplateIssueKind.UnsupportedMailMergeControlField &&
                    issue.Name == "SKIPIF" &&
                    issue.Message.Contains("SKIPIF", StringComparison.Ordinal));

                string json = report.ToJson();
                Assert.Contains("\"kind\": \"UnsupportedMailMergeControlField\"", json);
                Assert.Contains("\"name\": \"NEXTIF\"", json);
                Assert.Contains("\"name\": \"SKIPIF\"", json);

                string markdown = report.ToMarkdown();
                Assert.Contains("UnsupportedMailMergeControlField", markdown);
                Assert.Contains("Word-native mail-merge record-control field", markdown);
            }
        }

        private static SimpleField CreatePreflightMergeField(string name) {
            return new SimpleField(new Run(new Text("Placeholder"))) {
                Instruction = " MERGEFIELD  \"" + name + "\" "
            };
        }

        private static Paragraph CreatePreflightComplexFieldParagraph(string instruction, string result) {
            return new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = instruction }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(result)),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        private static void ReplaceTableCellPreflightContent(TableCell cell, params Paragraph[] paragraphs) {
            cell.RemoveAllChildren<Paragraph>();
            foreach (Paragraph paragraph in paragraphs) {
                cell.Append(paragraph);
            }
        }
    }
}
