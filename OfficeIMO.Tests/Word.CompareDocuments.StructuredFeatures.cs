using System;
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
        public void CompareStructureReportsFieldInstructionAndResultChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_field_diff.docx");
            CreateDocumentWithSimpleField(sourcePath, " AUTHOR ", "Alice");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_field_diff.docx");
            CreateDocumentWithSimpleField(targetPath, " TITLE ", "Quarterly report");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding field = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Field &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("field[0]", field.Location);
            Assert.Contains("Body", field.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("/word/document.xml", field.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("field[0]", field.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("AUTHOR", field.SourceText, StringComparison.Ordinal);
            Assert.Contains("Alice", field.SourceText, StringComparison.Ordinal);
            Assert.Contains("TITLE", field.TargetText, StringComparison.Ordinal);
            Assert.Contains("Quarterly report", field.TargetText, StringComparison.Ordinal);
        }

        [Fact]
        public void CompareStructureReportsContentControlAliasTagAndBindingChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_content_control_diff.docx");
            CreateDocumentWithBoundContentControl(
                sourcePath,
                alias: "Client",
                tag: "Client.Name",
                storeItemId: "{11111111-1111-1111-1111-111111111111}",
                xpath: "/root/client/name",
                text: "Contoso");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_content_control_diff.docx");
            CreateDocumentWithBoundContentControl(
                targetPath,
                alias: "Customer",
                tag: "Customer.Name",
                storeItemId: "{22222222-2222-2222-2222-222222222222}",
                xpath: "/root/customer/name",
                text: "Contoso");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding contentControl = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.ContentControl &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("content-control[0]", contentControl.Location);
            Assert.Contains("Body", contentControl.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("/word/document.xml", contentControl.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("content-control[0]", contentControl.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("alias=Client", contentControl.SourceText, StringComparison.Ordinal);
            Assert.Contains("tag=Client.Name", contentControl.SourceText, StringComparison.Ordinal);
            Assert.Contains("/root/client/name", contentControl.SourceText, StringComparison.Ordinal);
            Assert.Contains("alias=Customer", contentControl.TargetText, StringComparison.Ordinal);
            Assert.Contains("tag=Customer.Name", contentControl.TargetText, StringComparison.Ordinal);
            Assert.Contains("/root/customer/name", contentControl.TargetText, StringComparison.Ordinal);
        }

        [Fact]
        public void CompareStructureReturnsRichFieldAndContentControlFindingsInStableOrder() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_feature_order.docx");
            CreateDocumentWithRichFeatureComparisonInputs(
                sourcePath,
                contentControlAlias: "Client",
                contentControlTag: "Client.Name",
                contentControlFieldInstruction: " MERGEFIELD ClientName ",
                contentControlFieldResult: "Contoso",
                tableControlAlias: "Approval",
                tableControlTag: "Approval.State",
                tableFieldInstruction: " DOCPROPERTY Status ",
                tableFieldResult: "Draft",
                headerFieldInstruction: " TITLE ",
                headerFieldResult: "Draft plan");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_feature_order.docx");
            CreateDocumentWithRichFeatureComparisonInputs(
                targetPath,
                contentControlAlias: "Customer",
                contentControlTag: "Customer.Name",
                contentControlFieldInstruction: " MERGEFIELD CustomerName ",
                contentControlFieldResult: "Fabrikam",
                tableControlAlias: "Decision",
                tableControlTag: "Decision.State",
                tableFieldInstruction: " DOCPROPERTY ApprovalStatus ",
                tableFieldResult: "Approved",
                headerFieldInstruction: " SUBJECT ",
                headerFieldResult: "Approved plan");

            var options = new WordComparisonOptions {
                IncludedScopes = new HashSet<WordComparisonScope> {
                    WordComparisonScope.Field,
                    WordComparisonScope.ContentControl
                }
            };

            WordComparisonResult firstResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath, options);
            WordComparisonResult secondResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath, options);

            string[] firstSequence = firstResult.Findings.Select(FormatFindingSequenceEntry).ToArray();
            string[] secondSequence = secondResult.Findings.Select(FormatFindingSequenceEntry).ToArray();

            Assert.Equal(firstSequence, secondSequence);
            Assert.Equal(new[] {
                "Field|Modified|field[0]",
                "Field|Modified|field[1]",
                "ContentControl|Modified|content-control[0]",
                "ContentControl|Modified|content-control[1]",
                "Field|Modified|field[2]"
            }, firstSequence);

            Assert.Contains(firstResult.Findings, finding =>
                finding.Scope == WordComparisonScope.Field &&
                finding.Location == "field[0]" &&
                finding.DetailedLocation.Contains("content-control", StringComparison.Ordinal));
            Assert.Contains(firstResult.Findings, finding =>
                finding.Scope == WordComparisonScope.Field &&
                finding.Location == "field[1]" &&
                finding.DetailedLocation.Contains("table", StringComparison.Ordinal));
            Assert.Contains(firstResult.Findings, finding =>
                finding.Scope == WordComparisonScope.ContentControl &&
                finding.Location == "content-control[1]" &&
                finding.DetailedLocation.Contains("table", StringComparison.Ordinal));
            Assert.Contains(firstResult.Findings, finding =>
                finding.Scope == WordComparisonScope.Field &&
                finding.Location == "field[2]" &&
                finding.DetailedLocation.Contains("Header", StringComparison.Ordinal));
        }

        private static void CreateDocumentWithSimpleField(string path, string instruction, string resultText) {
            using WordDocument document = WordDocument.Create(path);
            document.AddParagraph("Placeholder");
            document.Save(false);

            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, true);
            Body body = wordDocument.MainDocumentPart!.Document.Body!;
            body.RemoveAllChildren<Paragraph>();
            body.PrependChild(new Paragraph(
                new Run(new Text("Field: ")),
                new SimpleField(
                    new Run(new Text(resultText) { Space = SpaceProcessingModeValues.Preserve })) {
                    Instruction = instruction
                }));
            wordDocument.MainDocumentPart.Document.Save();
        }

        private static void CreateDocumentWithRichFeatureComparisonInputs(
            string path,
            string contentControlAlias,
            string contentControlTag,
            string contentControlFieldInstruction,
            string contentControlFieldResult,
            string tableControlAlias,
            string tableControlTag,
            string tableFieldInstruction,
            string tableFieldResult,
            string headerFieldInstruction,
            string headerFieldResult) {
            using WordDocument document = WordDocument.Create(path);
            document.AddParagraph("Opening");
            document._document.Body!.Append(new SdtBlock(
                new SdtProperties(
                    new SdtAlias { Val = contentControlAlias },
                    new Tag { Val = contentControlTag }),
                new SdtContentBlock(
                    new Paragraph(
                        new Run(new Text("Account: ") { Space = SpaceProcessingModeValues.Preserve }),
                        CreateComparisonSimpleField(contentControlFieldInstruction, contentControlFieldResult)))));

            WordTable table = document.AddTable(1, 2);
            table.Rows[0].Cells[0]._tableCell.RemoveAllChildren<Paragraph>();
            table.Rows[0].Cells[0]._tableCell.Append(new SdtBlock(
                new SdtProperties(
                    new SdtAlias { Val = tableControlAlias },
                    new Tag { Val = tableControlTag }),
                new SdtContentBlock(
                    new Paragraph(
                        new Run(new Text("Ready"))))));
            table.Rows[0].Cells[1].Paragraphs[0]._paragraph.Append(CreateComparisonSimpleField(tableFieldInstruction, tableFieldResult));

            document.AddHeadersAndFooters();
            document.Header.Default!._header.Append(new Paragraph(
                new Run(new Text("Header: ") { Space = SpaceProcessingModeValues.Preserve }),
                CreateComparisonSimpleField(headerFieldInstruction, headerFieldResult)));
            document.Save(false);
        }

        private static SimpleField CreateComparisonSimpleField(string instruction, string resultText) {
            return new SimpleField(
                new Run(
                    new Text(resultText) { Space = SpaceProcessingModeValues.Preserve })) {
                Instruction = instruction
            };
        }

        private static string FormatFindingSequenceEntry(WordComparisonFinding finding) {
            return string.Join("|", finding.Scope, finding.ChangeKind, finding.Location);
        }

        private static void CreateDocumentWithBoundContentControl(string path, string alias, string tag, string storeItemId, string xpath, string text) {
            using WordDocument document = WordDocument.Create(path);
            document.AddParagraph("Placeholder");
            document.Save(false);

            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, true);
            Body body = wordDocument.MainDocumentPart!.Document.Body!;
            body.RemoveAllChildren<Paragraph>();
            body.PrependChild(new SdtBlock(
                new SdtProperties(
                    new SdtAlias { Val = alias },
                    new Tag { Val = tag },
                    new DataBinding {
                        StoreItemId = storeItemId,
                        XPath = xpath,
                        PrefixMappings = "xmlns:ns0='urn:test'"
                    }),
                new SdtContentBlock(
                    new Paragraph(
                        new Run(
                            new Text(text))))));
            wordDocument.MainDocumentPart.Document.Save();
        }
    }
}
