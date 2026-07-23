using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OfficeIMO.Word;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CompareStructureAlignsInsertedTablesBeforeModifiedTables() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_inserted_table_before_modified.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable terms = doc.AddTable(1, 1);
                terms.Rows[0].Cells[0].Paragraphs[0].SetText("Terms");
                WordTable closing = doc.AddTable(1, 1);
                closing.Rows[0].Cells[0].Paragraphs[0].SetText("Closing");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_inserted_table_before_modified.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable summary = doc.AddTable(1, 1);
                summary.Rows[0].Cells[0].Paragraphs[0].SetText("Summary");
                WordTable terms = doc.AddTable(1, 1);
                terms.Rows[0].Cells[0].Paragraphs[0].SetText("Terms updated");
                WordTable closing = doc.AddTable(1, 1);
                closing.Rows[0].Cells[0].Paragraphs[0].SetText("Closing");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Table &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Summary");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Terms" &&
                finding.TargetText == "Terms updated");
            Assert.DoesNotContain(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Terms" &&
                finding.TargetText == "Summary");
        }

        [Fact]
        public void CompareStructureAlignsInsertedCellsBeforeModifiedCells() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_inserted_cell_before_modified.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 2);
                SetTableTexts(table, "Owner", "Closing");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_inserted_cell_before_modified.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 3);
                SetTableTexts(table, "Evidence", "Owner updated", "Closing");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Evidence");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Owner" &&
                finding.TargetText == "Owner updated");
        }

        [Fact]
        public void CompareStructureAlignsInsertedImagesBeforeModifiedImages() {
            string logoPath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string replacementPath = Path.Combine(_directoryWithImages, "Kulek.jpg");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_inserted_image_before_modified.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph().AddImage(logoPath, 80, 40);
                doc.AddParagraph().AddImage(replacementPath, 50, 50);
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_inserted_image_before_modified.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph().AddImage(replacementPath, 120, 60);
                doc.AddParagraph().AddImage(replacementPath, 80, 40);
                doc.AddParagraph().AddImage(replacementPath, 50, 50);
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding inserted = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted);
            WordComparisonFinding modified = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("image[0]", inserted.Location);
            Assert.Equal("image[1]", modified.Location);
        }

        [Fact]
        public void CompareStructureDistinguishesDefaultAndFirstPageHeaders() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_default_header_variant.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.HeaderDefaultOrCreate.AddParagraph("Classification: Confidential");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_first_header_variant.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.HeaderFirstOrCreate.AddParagraph("Classification: Confidential");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted &&
                finding.SourceText == "Classification: Confidential");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Classification: Confidential");
        }

        [Fact]
        public void CompareStructurePreservesMovedFootnoteReferenceMarkers() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_moved_footnote_marker.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Policy").AddFootNote("Policy note");
                doc.AddParagraph("Policy");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_moved_footnote_marker.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Policy");
                doc.AddParagraph("Policy").AddFootNote("Policy note");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                ((finding.SourceText?.Contains("[FootnoteReference:", System.StringComparison.Ordinal) ?? false) ||
                 (finding.TargetText?.Contains("[FootnoteReference:", System.StringComparison.Ordinal) ?? false)));
        }

        [Fact]
        public void CompareStructureTraversesTablesAndImagesInsideNotes() {
            string logoPath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string replacementPath = Path.Combine(_directoryWithImages, "Kulek.jpg");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_note_tables_images.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Footnote anchor").AddFootNote("Footnote body");
                doc.AddParagraph("Endnote anchor").AddEndNote("Endnote body");
                doc.Save();
            }

            AppendTableAndImageToFirstFootnote(sourcePath, "Footnote table pending", logoPath);
            AppendTableAndImageToFirstEndnote(sourcePath, "Endnote table pending", logoPath);

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_note_tables_images.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Footnote anchor").AddFootNote("Footnote body");
                doc.AddParagraph("Endnote anchor").AddEndNote("Endnote body");
                doc.Save();
            }

            AppendTableAndImageToFirstFootnote(targetPath, "Footnote table approved", replacementPath);
            AppendTableAndImageToFirstEndnote(targetPath, "Endnote table approved", replacementPath);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Footnote table pending" &&
                finding.TargetText == "Footnote table approved");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Endnote table pending" &&
                finding.TargetText == "Endnote table approved");
            Assert.True(result.Findings.Count(finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified) >= 2);
        }

        [Fact]
        public void CompareStructureKeysNoteTableSnapshotsByStableNoteIds() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_note_table_stable_id.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Stable footnote anchor").AddFootNote("Stable footnote body");
                doc.Save();
            }

            SetReferencedFootnoteIds(sourcePath, 10);
            AppendTableToFootnote(sourcePath, 10, "Stable footnote table");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_note_table_stable_id.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Inserted footnote anchor").AddFootNote("Inserted footnote body");
                doc.AddParagraph("Stable footnote anchor").AddFootNote("Stable footnote body");
                doc.Save();
            }

            SetReferencedFootnoteIds(targetPath, 9, 10);
            AppendTableToFootnote(targetPath, 10, "Stable footnote table");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.DoesNotContain(result.Findings, finding =>
                finding.Scope is WordComparisonScope.Table or WordComparisonScope.TableCell);
        }

        [Fact]
        public void CompareStructureSplitsChangedParagraphMovedAcrossParts() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_changed_body_to_header.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Status: Pending");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_changed_body_to_header.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.HeaderDefaultOrCreate.AddParagraph("Status: Approved");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted &&
                finding.SourceText == "Status: Pending");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Status: Approved");
            Assert.DoesNotContain(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Status: Pending" &&
                finding.TargetText == "Status: Approved");
        }

        [Fact]
        public void CompareStructureIgnoresPictureLocalDrawingIds() {
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_picture_local_ids.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph().AddImage(imagePath, 80, 40);
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_picture_local_ids.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph().AddImage(imagePath, 80, 40);
                doc.Save();
            }

            MutatePictureLocalDrawingIds(targetPath, 4242U, "Different local picture id");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.DoesNotContain(result.Findings, finding => finding.Scope == WordComparisonScope.Image);
        }

        [Fact]
        public void CompareStructureAlignsMultipleInsertedParagraphsBeforeModifiedParagraph() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_multiple_inserted_paragraphs_before_modified.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Status: Pending");
                doc.AddParagraph("Closing");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_multiple_inserted_paragraphs_before_modified.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Evidence row one");
                doc.AddParagraph("Evidence row two");
                doc.AddParagraph("Status: Approved");
                doc.AddParagraph("Closing");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Evidence row one");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Evidence row two");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Status: Pending" &&
                finding.TargetText == "Status: Approved");
            Assert.DoesNotContain(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Status: Pending" &&
                finding.TargetText == "Evidence row one");
        }

        [Fact]
        public void CompareStructureAlignsModifiedParagraphInMiddleOfLargeInsertionRange() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_large_inserted_paragraph_range.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Status: Pending");
                doc.AddParagraph("Closing");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_large_inserted_paragraph_range.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                for (int index = 0; index < 300; index++) {
                    doc.AddParagraph("Evidence before " + index);
                }

                doc.AddParagraph("Status: Approved");
                for (int index = 0; index < 300; index++) {
                    doc.AddParagraph("Evidence after " + index);
                }

                doc.AddParagraph("Closing");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Equal(600, result.Findings.Count(finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText?.StartsWith("Evidence ", StringComparison.Ordinal) == true));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Status: Pending" &&
                finding.TargetText == "Status: Approved");
            Assert.DoesNotContain(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Status: Pending" &&
                finding.TargetText?.StartsWith("Evidence ", StringComparison.Ordinal) == true);
        }

        [Fact]
        public void CompareStructureRetainsModifiedParagraphWithStrongInternalMatch() {
            string sharedMiddle = string.Concat(Enumerable.Repeat("shared middle content ", 12));
            string sourceText = "A source start " + sharedMiddle + " source end";
            string targetText = "B target start " + sharedMiddle + " target end";
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_internal_alignment.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph(sourceText);
                doc.AddParagraph("Closing");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_internal_alignment.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                for (int index = 0; index < 260; index++) {
                    doc.AddParagraph("A inserted candidate " + index);
                }

                doc.AddParagraph(targetText);
                doc.AddParagraph("Closing");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == sourceText &&
                finding.TargetText == targetText);
            Assert.DoesNotContain(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == sourceText &&
                finding.TargetText?.StartsWith("A inserted candidate ", StringComparison.Ordinal) == true);
        }

        [Fact]
        public void CompareStructureReadsExplicitNormalFootnotes() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_explicit_normal_footnote.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Policy").AddFootNote("Legal note pending");
                doc.Save();
            }

            MarkFirstFootnoteAsExplicitNormal(sourcePath);

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_explicit_normal_footnote.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Policy").AddFootNote("Legal note approved");
                doc.Save();
            }

            MarkFirstFootnoteAsExplicitNormal(targetPath);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Legal note pending" &&
                finding.TargetText == "Legal note approved");
        }

        [Fact]
        public void CompareStructureSkipsUnreferencedHeaderParts() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_unref_header.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Body");
                doc.Save();
            }

            AddUnreferencedHeaderParagraph(sourcePath, "Dormant header");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_unref_header.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Body");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.DoesNotContain(result.Findings, finding =>
                (finding.SourceText?.Contains("Dormant header", System.StringComparison.Ordinal) ?? false) ||
                (finding.TargetText?.Contains("Dormant header", System.StringComparison.Ordinal) ?? false));
        }

        [Fact]
        public void CompareStructurePreservesSymbolRuns() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_symbol_run.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("placeholder");
                doc.Save();
            }

            ReplaceBodyParagraph(sourcePath, CreateSymbolParagraph("Choice ", "Wingdings", "F0FC"));

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_symbol_run.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("placeholder");
                doc.Save();
            }

            ReplaceBodyParagraph(targetPath, CreateSymbolParagraph("Choice ", "Wingdings", "F0A3"));

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                (finding.SourceText?.Contains("[Symbol:Wingdings:F0FC]", System.StringComparison.Ordinal) ?? false) &&
                (finding.TargetText?.Contains("[Symbol:Wingdings:F0A3]", System.StringComparison.Ordinal) ?? false));
        }

        [Fact]
        public void CompareStructureDistinguishesLiteralBreakMarkerTextFromRealBreaks() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_literal_page_break_marker.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Before[PageBreak]After");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_real_page_break_marker.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("placeholder");
                doc.Save();
            }

            ReplaceBodyParagraph(targetPath, new Paragraph(
                new Run(new Text("Before")),
                new Run(new Break { Type = BreakValues.Page }),
                new Run(new Text("After"))));

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Before[PageBreak]After" &&
                finding.TargetText == "Before[PageBreak]After");
        }

        private static void AppendTableAndImageToFirstFootnote(string path, string tableText, string imagePath) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Footnote footnote = document.MainDocumentPart!.FootnotesPart!.Footnotes!.Elements<Footnote>().First(item => item.Type == null);
            footnote.Append(CreateOneCellTable(tableText));
            AppendNoteImage(document.MainDocumentPart.FootnotesPart, footnote, imagePath);
            document.MainDocumentPart.FootnotesPart.Footnotes.Save();
        }

        private static void AppendTableToFootnote(string path, long footnoteId, string tableText) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Footnote footnote = document.MainDocumentPart!.FootnotesPart!.Footnotes!
                .Elements<Footnote>()
                .First(item => item.Id?.Value == footnoteId);
            footnote.Append(CreateOneCellTable(tableText));
            document.MainDocumentPart.FootnotesPart.Footnotes.Save();
        }

        private static void AppendParagraphToFootnote(string path, long footnoteId, string text) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Footnote footnote = document.MainDocumentPart!.FootnotesPart!.Footnotes!
                .Elements<Footnote>()
                .First(item => item.Id?.Value == footnoteId);
            footnote.Append(new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
            document.MainDocumentPart.FootnotesPart.Footnotes.Save();
        }

        private static void SetReferencedFootnoteIds(string path, params long[] ids) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            FootnoteReference[] references = document.MainDocumentPart!.Document!.Body!.Descendants<FootnoteReference>().ToArray();
            Footnote[] footnotes = document.MainDocumentPart.FootnotesPart!.Footnotes!
                .Elements<Footnote>()
                .Where(item => item.Type == null || item.Type.Value == FootnoteEndnoteValues.Normal)
                .ToArray();

            Assert.True(ids.Length <= references.Length, "Not enough footnote references to assign stable ids.");
            Assert.True(ids.Length <= footnotes.Length, "Not enough footnotes to assign stable ids.");
            for (int index = 0; index < ids.Length; index++) {
                references[index].Id = ids[index];
                footnotes[index].Id = ids[index];
            }

            document.MainDocumentPart.Document.Save();
            document.MainDocumentPart.FootnotesPart.Footnotes.Save();
        }

        private static void AppendTableAndImageToFirstEndnote(string path, string tableText, string imagePath) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Endnote endnote = document.MainDocumentPart!.EndnotesPart!.Endnotes!.Elements<Endnote>().First(item => item.Type == null);
            endnote.Append(CreateOneCellTable(tableText));
            AppendNoteImage(document.MainDocumentPart.EndnotesPart, endnote, imagePath);
            document.MainDocumentPart.EndnotesPart.Endnotes.Save();
        }

        private static Table CreateOneCellTable(string text) {
            return new Table(
                new TableRow(
                    new TableCell(
                        new Paragraph(
                            new Run(
                                new Text(text))))));
        }

        private static void AppendNoteImage(FootnotesPart part, OpenXmlElement container, string imagePath) {
            ImagePart imagePart = Path.GetExtension(imagePath).Equals(".jpg", System.StringComparison.OrdinalIgnoreCase) ||
                                  Path.GetExtension(imagePath).Equals(".jpeg", System.StringComparison.OrdinalIgnoreCase)
                ? part.AddImagePart(ImagePartType.Jpeg)
                : part.AddImagePart(ImagePartType.Png);
            FeedNoteImage(part, imagePart, container, imagePath);
        }

        private static void AppendNoteImage(EndnotesPart part, OpenXmlElement container, string imagePath) {
            ImagePart imagePart = Path.GetExtension(imagePath).Equals(".jpg", System.StringComparison.OrdinalIgnoreCase) ||
                                  Path.GetExtension(imagePath).Equals(".jpeg", System.StringComparison.OrdinalIgnoreCase)
                ? part.AddImagePart(ImagePartType.Jpeg)
                : part.AddImagePart(ImagePartType.Png);
            FeedNoteImage(part, imagePart, container, imagePath);
        }

        private static void FeedNoteImage(OpenXmlPartContainer part, ImagePart imagePart, OpenXmlElement container, string imagePath) {
            using (FileStream stream = File.OpenRead(imagePath)) {
                imagePart.FeedData(stream);
            }

            string relationshipId = part.GetIdOfPart(imagePart);
            container.Append(new Paragraph(new Run(CreateInlineDrawing(relationshipId, 914400L, 457200L, 1U, "Note image"))));
        }

        private static DocumentFormat.OpenXml.Wordprocessing.Drawing CreateInlineDrawing(string relationshipId, long widthEmu, long heightEmu, uint id, string name) {
            return new DocumentFormat.OpenXml.Wordprocessing.Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties { Id = id, Name = name },
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties { Id = id, Name = name },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip { Embed = relationshipId },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
                { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U });
        }

        private static void MutatePictureLocalDrawingIds(string path, uint id, string name) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            foreach (PIC.NonVisualDrawingProperties properties in document.MainDocumentPart!.Document.Descendants<PIC.NonVisualDrawingProperties>()) {
                properties.Id = id;
                properties.Name = name;
            }

            document.MainDocumentPart.Document.Save();
        }

        private static void MarkFirstFootnoteAsExplicitNormal(string path) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Footnote footnote = document.MainDocumentPart!.FootnotesPart!.Footnotes!.Elements<Footnote>().First(item => item.Type == null);
            footnote.Type = FootnoteEndnoteValues.Normal;
            document.MainDocumentPart.FootnotesPart.Footnotes.Save();
        }

        private static void AddUnreferencedHeaderParagraph(string path, string text) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            HeaderPart headerPart = document.MainDocumentPart!.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new Paragraph(new Run(new Text(text))));
            headerPart.Header.Save();
        }

        private static Paragraph CreateSymbolParagraph(string prefix, string font, string symbolChar) {
            return new Paragraph(
                new Run(new Text(prefix)),
                new Run(new SymbolChar { Font = font, Char = symbolChar }));
        }

        private static void ReplaceBodyParagraph(string path, Paragraph paragraph) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Body body = document.MainDocumentPart!.Document.Body!;
            body.RemoveAllChildren();
            body.Append(paragraph);
            document.MainDocumentPart.Document.Save();
        }
    }
}
