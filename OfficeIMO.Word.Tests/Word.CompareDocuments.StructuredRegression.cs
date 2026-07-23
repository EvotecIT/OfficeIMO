using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CompareStructureTreatsMovedBodyParagraphToHeaderAsDeleteInsert() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_body_to_header.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Body anchor");
                doc.AddParagraph("Classification: Confidential");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_body_to_header.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Body anchor");
                doc.AddHeadersAndFooters();
                doc.Header.Default!.AddParagraph("Classification: Confidential");
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
            Assert.DoesNotContain(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Classification: Confidential" &&
                finding.TargetText == "Classification: Confidential");
        }

        [Fact]
        public void CompareStructurePreservesPageAndColumnBreaksInParagraphText() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_page_break.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                AddBreakParagraph(doc, BreakValues.Page);
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_column_break.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                AddBreakParagraph(doc, BreakValues.Column);
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding modified = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("Before[PageBreak]After", modified.SourceText);
            Assert.Equal("Before[ColumnBreak]After", modified.TargetText);
        }

        [Fact]
        public void CompareStructureReportsVmlImageReplacement() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_vml_image.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                AddVmlImageParagraph(doc, Path.Combine(_directoryWithImages, "EvotecLogo.png"));
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_vml_image.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                AddVmlImageParagraph(doc, Path.Combine(_directoryWithImages, "Kulek.jpg"));
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding image = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("image[0]", image.Location);
        }

        [Fact]
        public void CompareStructureReadsBlockContentControlTextInsideTableCells() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_cell_sdt.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 1);
                ReplaceCellWithBlockContentControl(table.Rows[0].Cells[0], "Evidence pending");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_cell_sdt.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 1);
                ReplaceCellWithBlockContentControl(table.Rows[0].Cells[0], "Evidence approved");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding cell = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("Evidence pending", cell.SourceText);
            Assert.Equal("Evidence approved", cell.TargetText);
        }

        [Fact]
        public void CompareStructureAlignsInsertionBeforeModifiedParagraph() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_insert_before_modified.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Terms and conditions apply.");
                doc.AddParagraph("Closing section");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_insert_before_modified.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("New executive summary");
                doc.AddParagraph("Terms and conditions apply after approval.");
                doc.AddParagraph("Closing section");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Equal(2, result.Findings.Count);
            WordComparisonFinding inserted = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted);
            Assert.Equal("paragraph[0]", inserted.Location);
            Assert.Equal("New executive summary", inserted.TargetText);

            WordComparisonFinding modified = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("paragraph[1]", modified.Location);
            Assert.Equal("Terms and conditions apply.", modified.SourceText);
            Assert.Equal("Terms and conditions apply after approval.", modified.TargetText);
        }

        [Fact]
        public void CompareStructureReportsImageGeometryChangesForSamePayload() {
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_image_geometry.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph().AddImage(imagePath, 40, 40);
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_image_geometry.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph().AddImage(imagePath, 120, 40);
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding image = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("Image layout changed.", image.Message);
        }

        [Fact]
        public void CompareStructureHandlesLargeParagraphSetsWithBoundedAlignment() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_large_alignment.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                AddNumberedParagraphs(doc, 1000);
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_large_alignment.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Inserted cover note");
                AddNumberedParagraphs(doc, 1000);
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding inserted = Assert.Single(result.Findings);
            Assert.Equal(WordComparisonScope.Paragraph, inserted.Scope);
            Assert.Equal(WordComparisonChangeKind.Inserted, inserted.ChangeKind);
            Assert.Equal("Inserted cover note", inserted.TargetText);
        }

        [Fact]
        public void CompareStructureHandlesLargeSimilarParagraphsWithBoundedSimilarity() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_large_paragraph_similarity.docx");
            string sourceText = new string('A', 50000);
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph(sourceText);
                doc.AddParagraph("Closing");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_large_paragraph_similarity.docx");
            string targetText = new string('A', 49999) + "B";
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Inserted cover note");
                doc.AddParagraph(targetText);
                doc.AddParagraph("Closing");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Inserted cover note");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == sourceText &&
                finding.TargetText == targetText);
        }

        [Fact]
        public void CompareStructureReturnsMixedFindingsInDocumentOrder() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_mixed_order.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Opening");
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Risk: Open");
                doc.AddParagraph("Decision: Draft");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_mixed_order.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Opening");
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Risk: Closed");
                doc.AddParagraph("Decision: Approved");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Equal(2, result.Findings.Count);
            Assert.Equal(WordComparisonScope.TableCell, result.Findings[0].Scope);
            Assert.Equal("Risk: Open", result.Findings[0].SourceText);
            Assert.Equal(WordComparisonScope.Paragraph, result.Findings[1].Scope);
            Assert.Equal("Decision: Draft", result.Findings[1].SourceText);
        }

        [Fact]
        public void CompareStructureReturnsTableCellFindingsInRowMajorOrder() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_table_cell_order.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(2, 3);
                SetTableTexts(table, "A1", "A2", "A3", "B1", "B2", "B3");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_table_cell_order.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(2, 3);
                SetTableTexts(table, "A1", "A2", "A3 changed", "B1 changed", "B2", "B3");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);
            WordComparisonFinding[] cellFindings = result.Findings
                .Where(finding => finding.Scope == WordComparisonScope.TableCell)
                .ToArray();

            Assert.Equal(2, cellFindings.Length);
            Assert.Equal("table[0]/row[0]/cell[2]", cellFindings[0].Location);
            Assert.Equal("table[0]/row[1]/cell[0]", cellFindings[1].Location);
        }

        [Fact]
        public void CompareStructureKeepsTableChildFindingsBeforeFollowingParagraphs() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_table_before_paragraph.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(2, 1);
                SetTableTexts(table, "Owner", "Risk: Open");
                doc.AddParagraph("Decision: Draft");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_table_before_paragraph.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(2, 1);
                SetTableTexts(table, "Owner", "Risk: Closed");
                doc.AddParagraph("Decision: Approved");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Equal(2, result.Findings.Count);
            Assert.Equal(WordComparisonScope.TableCell, result.Findings[0].Scope);
            Assert.Equal("Risk: Open", result.Findings[0].SourceText);
            Assert.Equal(WordComparisonScope.Paragraph, result.Findings[1].Scope);
            Assert.Equal("Decision: Draft", result.Findings[1].SourceText);
        }

        [Fact]
        public void CompareStructureDisambiguatesCellBoundariesWhenTextContainsSeparator() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_cell_boundary_key.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 2);
                SetTableTexts(table, "A | B", "C");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_cell_boundary_key.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 2);
                SetTableTexts(table, "A", "B | C");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);
            WordComparisonFinding[] cellFindings = result.Findings
                .Where(finding => finding.Scope == WordComparisonScope.TableCell)
                .ToArray();

            Assert.Equal(2, cellFindings.Length);
            Assert.Equal("A | B", cellFindings[0].SourceText);
            Assert.Equal("A", cellFindings[0].TargetText);
            Assert.Equal("C", cellFindings[1].SourceText);
            Assert.Equal("B | C", cellFindings[1].TargetText);
        }

        [Fact]
        public void CompareStructureAlignsInsertedTableRowsBeforeModifiedRows() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_inserted_row_before_modified.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(2, 2);
                SetTableTexts(table, "MFA", "Security", "Logging", "Platform");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_inserted_row_before_modified.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(3, 2);
                SetTableTexts(table, "Review", "Compliance", "MFA", "Identity", "Logging", "Platform");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding inserted = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableRow &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted);
            WordComparisonFinding modified = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            WordComparisonFinding[] findings = result.Findings.ToArray();
            Assert.True(System.Array.IndexOf(findings, inserted) < System.Array.IndexOf(findings, modified));
            Assert.Equal("Review | Compliance", inserted.TargetText);
            Assert.Equal("Security", modified.SourceText);
            Assert.Equal("Identity", modified.TargetText);
        }

        [Fact]
        public void CompareStructureAlignsModifiedTableRowInMiddleOfLargeInsertionRange() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_large_inserted_row_range.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(2, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Status: Pending");
                table.Rows[1].Cells[0].Paragraphs[0].SetText("Closing");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_large_inserted_row_range.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(602, 1);
                for (int index = 0; index < 300; index++) {
                    table.Rows[index].Cells[0].Paragraphs[0].SetText("Evidence before " + index);
                }

                table.Rows[300].Cells[0].Paragraphs[0].SetText("Status: Approved");
                for (int index = 0; index < 300; index++) {
                    table.Rows[index + 301].Cells[0].Paragraphs[0].SetText("Evidence after " + index);
                }

                table.Rows[601].Cells[0].Paragraphs[0].SetText("Closing");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Equal(600, result.Findings.Count(finding =>
                finding.Scope == WordComparisonScope.TableRow &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText?.StartsWith("Evidence ", StringComparison.Ordinal) == true));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Status: Pending" &&
                finding.TargetText == "Status: Approved");
        }

        [Fact]
        public void CompareStructureReportsFootnoteAndEndnoteBodyChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_notes.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Footnote anchor").AddFootNote("Footnote pending");
                doc.AddParagraph("Endnote anchor").AddEndNote("Endnote pending");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_notes.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Footnote anchor").AddFootNote("Footnote approved");
                doc.AddParagraph("Endnote anchor").AddEndNote("Endnote approved");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Footnote pending" &&
                finding.TargetText == "Footnote approved");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Endnote pending" &&
                finding.TargetText == "Endnote approved");
        }

        [Fact]
        public void CompareStructureTreatsMovedBodyTableToHeaderAsDeleteInsert() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_body_table_to_header.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Review owner");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_body_table_to_header.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddHeadersAndFooters();
                WordTable table = doc.Header.Default!.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Review owner");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Table &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted &&
                finding.SourceText == "Review owner");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Table &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Review owner");
        }

        [Fact]
        public void CompareStructureReportsTableRowShapeChangesWithSameJoinedText() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_table_shape.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 2);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Owner");
                table.Rows[0].Cells[1].Paragraphs[0].SetText("Status");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_table_shape.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Owner | Status");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableRow &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted &&
                finding.SourceText == "Owner | Status");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableRow &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Owner | Status");
        }

        [Fact]
        public void CompareStructureUsesInvariantCellParagraphSeparators() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_cell_paragraph_separator.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0]._tableCell.RemoveAllChildren<Paragraph>();
                table.Rows[0].Cells[0]._tableCell.Append(
                    new Paragraph(new Run(new Text("A"))),
                    new Paragraph(new Run(new Text("B"))));
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_cell_paragraph_separator.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0]._tableCell.RemoveAllChildren<Paragraph>();
                table.Rows[0].Cells[0]._tableCell.Append(
                    new Paragraph(
                        new Run(new Text("A")),
                        new Run(new Break()),
                        new Run(new Text("B"))));
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding cell = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("A[ParagraphBreak]B", cell.SourceText);
            Assert.Equal("A\nB", cell.TargetText);
        }

        [Fact]
        public void CompareStructureTreatsMovedBodyImageToHeaderAsDeleteInsert() {
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_body_image_to_header.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph().AddImage(imagePath, 80, 40);
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_body_image_to_header.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddHeadersAndFooters();
                doc.Header.Default!.AddParagraph().AddImage(imagePath, 80, 40);
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted);
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted);
            Assert.DoesNotContain(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
        }

        [Fact]
        public void CompareStructureReportsExternalVmlImageUriChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_external_vml_image.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                AddExternalVmlImageParagraph(doc, "https://example.com/source.png");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_external_vml_image.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                AddExternalVmlImageParagraph(doc, "https://example.com/target.png");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding image = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("[Image: https://example.com/source.png]", image.SourceText);
            Assert.Equal("[Image: https://example.com/target.png]", image.TargetText);
        }

        [Fact]
        public void CompareStructureKeepsMixedDrawingAndVmlImagesInDocumentOrder() {
            string logoPath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string replacementPath = Path.Combine(_directoryWithImages, "Kulek.jpg");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_mixed_image_order.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                AddVmlImageParagraph(doc, logoPath);
                doc.AddParagraph().AddImage(logoPath, 80, 40);
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_mixed_image_order.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                AddVmlImageParagraph(doc, replacementPath);
                doc.AddParagraph().AddImage(logoPath, 80, 40);
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding image = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("image[0]", image.Location);
        }

        [Fact]
        public void CompareStructureReportsTextBoxParagraphChangesOnce() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_textbox.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                AddTextBoxParagraph(doc, "Callout pending");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_textbox.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                AddTextBoxParagraph(doc, "Callout approved");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);
            WordComparisonFinding[] paragraphFindings = result.Findings
                .Where(finding => finding.Scope == WordComparisonScope.Paragraph)
                .ToArray();

            WordComparisonFinding paragraph = Assert.Single(paragraphFindings);
            Assert.Equal("Callout pending", paragraph.SourceText);
            Assert.Equal("Callout approved", paragraph.TargetText);
        }

        private static void AddBreakParagraph(WordDocument document, BreakValues breakType) {
            WordParagraph paragraph = document.AddParagraph("Before");
            paragraph._paragraph.Append(new Run(new Break { Type = breakType }));
            paragraph._paragraph.Append(new Run(new Text("After")));
        }

        private static void SetTableTexts(WordTable table, params string[] values) {
            int valueIndex = 0;
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    cell.Paragraphs[0].SetText(values[valueIndex]);
                    valueIndex++;
                }
            }
        }

        private static void AddNumberedParagraphs(WordDocument document, int count) {
            for (int index = 0; index < count; index++) {
                document.AddParagraph("Clause " + index.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
        }

        private static void AddVmlImageParagraph(WordDocument document, string imagePath) {
            MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;
            ImagePart imagePart = Path.GetExtension(imagePath).Equals(".jpg", System.StringComparison.OrdinalIgnoreCase) ||
                                  Path.GetExtension(imagePath).Equals(".jpeg", System.StringComparison.OrdinalIgnoreCase)
                ? mainPart.AddImagePart(ImagePartType.Jpeg)
                : mainPart.AddImagePart(ImagePartType.Png);
            using (FileStream stream = File.OpenRead(imagePath)) {
                imagePart.FeedData(stream);
            }

            string relationshipId = mainPart.GetIdOfPart(imagePart);
            var imageData = new V.ImageData {
                RelationshipId = relationshipId,
                Title = "Legacy image"
            };
            var shape = new V.Shape(imageData) {
                Id = "LegacyImage",
                Type = "#_x0000_t75",
                Style = "width:72pt;height:72pt",
                Filled = false,
                Stroked = false
            };

            document._document.Body!.Append(new Paragraph(new Run(new Picture(shape))));
        }

        private static void AddTextBoxParagraph(WordDocument document, string text) {
            var textBox = new V.TextBox(
                new TextBoxContent(
                    new Paragraph(
                        new Run(
                            new Text(text)))));
            var shape = new V.Shape(textBox) {
                Id = "Callout",
                Type = "#_x0000_t202",
                Style = "width:120pt;height:40pt",
                Filled = false,
                Stroked = true
            };

            document._document.Body!.Append(
                new Paragraph(
                    new Run(new Text("Host paragraph")),
                    new Run(new Picture(shape))));
        }

        private static void AddExternalVmlImageParagraph(WordDocument document, string uri) {
            MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;
            ExternalRelationship relationship = mainPart.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                new System.Uri(uri));
            var imageData = new V.ImageData {
                RelationshipId = relationship.Id,
                Title = "Linked legacy image"
            };
            var shape = new V.Shape(imageData) {
                Id = "LinkedLegacyImage",
                Type = "#_x0000_t75",
                Style = "width:72pt;height:72pt",
                Filled = false,
                Stroked = false
            };

            document._document.Body!.Append(new Paragraph(new Run(new Picture(shape))));
        }

        private static void ReplaceCellWithBlockContentControl(WordTableCell cell, string text) {
            cell._tableCell.RemoveAllChildren<Paragraph>();
            cell._tableCell.Append(new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "Evidence" }),
                new SdtContentBlock(new Paragraph(new Run(new Text(text))))));
        }
    }
}
