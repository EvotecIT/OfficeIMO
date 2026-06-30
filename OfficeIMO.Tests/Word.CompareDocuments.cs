using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CompareDetectsInsertedText() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_insert.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Hello");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_insert.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Hello World");
                doc.Save(false);
            }

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            Body body = result._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            InsertedRun? ins = body.Descendants<InsertedRun>().FirstOrDefault();
            Assert.NotNull(ins);
            Assert.Equal(" World", ins!.InnerText);
            AssertNoTempArtifact(result);
        }

        [Fact]
        public void CompareDetectsDeletedText() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_delete.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Hello World");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_delete.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Hello");
                doc.Save(false);
            }

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            Body body = result._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            DeletedRun? del = body.Descendants<DeletedRun>().FirstOrDefault();
            Assert.NotNull(del);
            Assert.Equal(" World", del!.InnerText);
            AssertNoTempArtifact(result);
        }

        [Fact]
        public void CompareDetectsFormattingChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_format.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Hello");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_format.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                var p = doc.AddParagraph("Hello");
                p.Bold = true;
                doc.Save(false);
            }

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            Body body = result._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Run run = body.Descendants<Run>().First();
            Assert.NotNull(run.RunProperties);
            Assert.NotNull(run.RunProperties!.RunPropertiesChange);
            AssertNoTempArtifact(result);
        }

        [Fact]
        public void CompareDetectsInsertedTableRow() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_row_insert.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Row1");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_row_insert.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(2, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Row1");
                table.Rows[1].Cells[0].Paragraphs[0].SetText("Row2");
                doc.Save(false);
            }

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            Body body = result._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            InsertedRun? ins = body.Descendants<InsertedRun>().FirstOrDefault(r => r.InnerText == "Row2");
            Assert.NotNull(ins);
            AssertNoTempArtifact(result);
        }

        [Fact]
        public void CompareDetectsDeletedTableRow() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_row_delete.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(2, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Row1");
                table.Rows[1].Cells[0].Paragraphs[0].SetText("Row2");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_row_delete.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Row1");
                doc.Save(false);
            }

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            Body body = result._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            DeletedRun? del = body.Descendants<DeletedRun>().FirstOrDefault(r => r.InnerText == "Row2");
            Assert.NotNull(del);
            AssertNoTempArtifact(result);
        }

        [Fact]
        public void CompareDetectsImageReplacement() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_image.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_image.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "snail.bmp"));
                doc.Save(false);
            }

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            Body body = result._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            InsertedRun? ins = body.Descendants<InsertedRun>().FirstOrDefault(r => r.InnerText == "[Image]");
            DeletedRun? del = body.Descendants<DeletedRun>().FirstOrDefault(r => r.InnerText == "[Image]");
            Assert.NotNull(ins);
            Assert.NotNull(del);
            AssertNoTempArtifact(result);
        }

        [Fact]
        public void CompareDetectsInsertedTableCellInExistingRow() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_cell_insert.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Left");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_cell_insert.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 2);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Left");
                table.Rows[0].Cells[1].Paragraphs[0].SetText("Right");
                doc.Save(false);
            }

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            Assert.Equal(2, result.Tables[0].Rows[0].CellsCount);

            Body body = result._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            InsertedRun? ins = body.Descendants<InsertedRun>().FirstOrDefault(r => r.InnerText == "Right");
            Assert.NotNull(ins);
            AssertNoTempArtifact(result);
        }

        [Fact]
        public void ComparePreservesListFormattingOnChangedParagraph() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_source_list_format.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                var list = doc.AddList(WordListStyle.Numbered);
                list.AddItem("Item 1");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_target_list_format.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                var list = doc.AddList(WordListStyle.Numbered);
                list.AddItem("Item 1 updated");
                doc.Save(false);
            }

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            WordParagraph? listParagraph = result.Paragraphs.FirstOrDefault(p => p.Text.Contains("Item 1"));

            Assert.NotNull(listParagraph);
            Assert.True(listParagraph!.IsListItem);
            AssertNoTempArtifact(result);
        }

        [Fact]
        public void CompareStructureReportsParagraphChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_paragraph.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Status: Draft");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_paragraph.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Status: Approved");
                doc.AddParagraph("Scope: Added");
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.True(result.HasChanges);
            WordComparisonFinding modified = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("paragraph[0]", modified.Location);
            Assert.Equal("Status: Draft", modified.SourceText);
            Assert.Equal("Status: Approved", modified.TargetText);

            WordComparisonFinding inserted = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted);
            Assert.Equal("paragraph[1]", inserted.Location);
            Assert.Equal("Scope: Added", inserted.TargetText);
        }

        [Fact]
        public void CompareStructureReportsTableCellAndRowChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_table.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 2);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Control");
                table.Rows[0].Cells[1].Paragraphs[0].SetText("Open");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_table.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Control");
                table.Rows[0].Cells[1].Paragraphs[0].SetText("Closed");
                table.Rows[1].Cells[0].Paragraphs[0].SetText("Evidence");
                table.Rows[1].Cells[1].Paragraphs[0].SetText("Added");
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding cell = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("table[0]/row[0]/cell[1]", cell.Location);
            Assert.Equal("Open", cell.SourceText);
            Assert.Equal("Closed", cell.TargetText);

            WordComparisonFinding row = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableRow &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted);
            Assert.Equal("table[0]/row[1]", row.Location);
            Assert.Equal("Evidence | Added", row.TargetText);
        }

        [Fact]
        public void CompareStructureReportsImageReplacement() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_image.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_image.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "snail.bmp"));
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding image = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("image[0]", image.Location);
            Assert.Equal("Image payload changed.", image.Message);
        }

        [Fact]
        public void CompareStructureAlignsParagraphInsertionsWithoutFalseShiftChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_paragraph_alignment.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Service Agreement").Style = WordParagraphStyles.Heading1;
                doc.AddParagraph("Payment is due within 30 days.");
                doc.AddParagraph("Audit logs must be retained for 90 days.");
                doc.AddParagraph("Both parties accept the terms.");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_paragraph_alignment.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Service Agreement").Style = WordParagraphStyles.Heading1;
                doc.AddParagraph("Payment is due within 14 days.");
                doc.AddParagraph("Late payment requires written escalation.");
                doc.AddParagraph("Audit logs must be retained for 90 days.");
                doc.AddParagraph("Both parties accept the terms.");
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Equal(2, result.Findings.Count);
            WordComparisonFinding modified = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("paragraph[1]", modified.Location);
            Assert.Equal("Payment is due within 30 days.", modified.SourceText);
            Assert.Equal("Payment is due within 14 days.", modified.TargetText);

            WordComparisonFinding inserted = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted);
            Assert.Equal("paragraph[2]", inserted.Location);
            Assert.Equal("Late payment requires written escalation.", inserted.TargetText);
            Assert.DoesNotContain(result.Findings, finding => finding.TargetText == "Audit logs must be retained for 90 days.");
        }

        [Fact]
        public void CompareStructureAlignsTableRowsAroundInsertedEvidenceRows() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_table_alignment.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(3, 2);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Control");
                table.Rows[0].Cells[1].Paragraphs[0].SetText("Owner");
                table.Rows[1].Cells[0].Paragraphs[0].SetText("MFA");
                table.Rows[1].Cells[1].Paragraphs[0].SetText("Security");
                table.Rows[2].Cells[0].Paragraphs[0].SetText("Logging");
                table.Rows[2].Cells[1].Paragraphs[0].SetText("Platform");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_table_alignment.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(4, 2);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Control");
                table.Rows[0].Cells[1].Paragraphs[0].SetText("Owner");
                table.Rows[1].Cells[0].Paragraphs[0].SetText("MFA");
                table.Rows[1].Cells[1].Paragraphs[0].SetText("Identity");
                table.Rows[2].Cells[0].Paragraphs[0].SetText("Review");
                table.Rows[2].Cells[1].Paragraphs[0].SetText("Compliance");
                table.Rows[3].Cells[0].Paragraphs[0].SetText("Logging");
                table.Rows[3].Cells[1].Paragraphs[0].SetText("Platform");
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Equal(2, result.Findings.Count);
            WordComparisonFinding cell = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("table[0]/row[1]/cell[1]", cell.Location);
            Assert.Equal("Security", cell.SourceText);
            Assert.Equal("Identity", cell.TargetText);

            WordComparisonFinding row = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableRow &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted);
            Assert.Equal("table[0]/row[2]", row.Location);
            Assert.Equal("Review | Compliance", row.TargetText);
            Assert.DoesNotContain(result.Findings, finding => finding.TargetText == "Logging | Platform");
        }

        [Fact]
        public void CompareStructureAlignsImagesAroundInsertedImages() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_image_alignment.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
                doc.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "snail.bmp"));
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_image_alignment.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
                doc.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"));
                doc.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "snail.bmp"));
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding image = Assert.Single(result.Findings);
            Assert.Equal(WordComparisonScope.Image, image.Scope);
            Assert.Equal(WordComparisonChangeKind.Inserted, image.ChangeKind);
            Assert.Equal("image[1]", image.Location);
            Assert.Equal(1, image.TargetIndex);
        }

        [Fact]
        public void CompareStructureTreatsFormattedRunsAsOneLogicalParagraph() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_formatted_runs.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Project Alpha approved.");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_formatted_runs.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordParagraph paragraph = doc.AddParagraph("Project ");
                paragraph.AddText("Alpha").Bold = true;
                paragraph.AddText(" approved.");
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(
                sourcePath,
                targetPath,
                new WordComparisonOptions {
                    CompareRunFormatting = false
                });

            Assert.Empty(result.Findings);
        }

        [Fact]
        public void CompareStructureAlignsTableCellsAroundInsertedCells() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_cell_alignment.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 3);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Control");
                table.Rows[0].Cells[1].Paragraphs[0].SetText("Owner");
                table.Rows[0].Cells[2].Paragraphs[0].SetText("Status");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_cell_alignment.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 4);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Control");
                table.Rows[0].Cells[1].Paragraphs[0].SetText("Evidence");
                table.Rows[0].Cells[2].Paragraphs[0].SetText("Owner");
                table.Rows[0].Cells[3].Paragraphs[0].SetText("Status");
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding cell = Assert.Single(result.Findings);
            Assert.Equal(WordComparisonScope.TableCell, cell.Scope);
            Assert.Equal(WordComparisonChangeKind.Inserted, cell.ChangeKind);
            Assert.Equal("table[0]/row[0]/cell[1]", cell.Location);
            Assert.Equal("Evidence", cell.TargetText);
        }

        [Fact]
        public void CompareStructureFindsNestedTableCellChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_nested_table.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 1);
                WordTable nested = table.Rows[0].Cells[0].AddTable(1, 1);
                nested.Rows[0].Cells[0].Paragraphs[0].SetText("Evidence pending");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_nested_table.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 1);
                WordTable nested = table.Rows[0].Cells[0].AddTable(1, 1);
                nested.Rows[0].Cells[0].Paragraphs[0].SetText("Evidence approved");
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding cell = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Evidence pending" &&
                finding.TargetText == "Evidence approved");
            Assert.Equal("table[1]/row[0]/cell[0]", cell.Location);
        }

        [Fact]
        public void CompareStructureAlignsTablesAroundInsertedTables() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_table_insert_alignment.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Existing terms");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_table_insert_alignment.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable inserted = doc.AddTable(1, 1);
                inserted.Rows[0].Cells[0].Paragraphs[0].SetText("New summary");
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Existing terms");
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding tableFinding = Assert.Single(result.Findings);
            Assert.Equal(WordComparisonScope.Table, tableFinding.Scope);
            Assert.Equal(WordComparisonChangeKind.Inserted, tableFinding.ChangeKind);
            Assert.Equal("table[0]", tableFinding.Location);
            Assert.Equal("New summary", tableFinding.TargetText);
        }

        [Fact]
        public void CompareStructurePreservesVisibleParagraphWhitespaceAndBlankParagraphs() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_visible_whitespace.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordParagraph paragraph = doc.AddParagraph("Column A");
                paragraph.AddTab();
                paragraph.AddText("Column B");
                doc.AddParagraph("Closing");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_visible_whitespace.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordParagraph paragraph = doc.AddParagraph("Column A");
                paragraph.AddBreak();
                paragraph.AddText("Column B");
                doc.AddParagraph();
                doc.AddParagraph("Closing");
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding modified = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("Column A\tColumn B", modified.SourceText);
            Assert.Equal("Column A\nColumn B", modified.TargetText);

            WordComparisonFinding inserted = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted);
            Assert.Equal("paragraph[1]", inserted.Location);
            Assert.Equal(string.Empty, inserted.TargetText);
        }

        [Fact]
        public void CompareStructureTreatsFormattedTableCellRunsAsOneLogicalParagraph() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_cell_formatted_runs.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Project Alpha approved");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_cell_formatted_runs.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 1);
                WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
                paragraph.SetText("Project ");
                paragraph.AddText("Alpha").Bold = true;
                paragraph.AddText(" approved");
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Empty(result.Findings);
        }

        [Fact]
        public void CompareStructureIncludesHeaderFooterParagraphsAndTables() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_header_footer.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddHeadersAndFooters();
                doc.Header.Default!.AddParagraph("Classification: Internal");
                WordTable footerTable = doc.Footer.Default!.AddTable(1, 2);
                footerTable.Rows[0].Cells[0].Paragraphs[0].SetText("Owner");
                footerTable.Rows[0].Cells[1].Paragraphs[0].SetText("Platform");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_header_footer.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddHeadersAndFooters();
                doc.Header.Default!.AddParagraph("Classification: Confidential");
                WordTable footerTable = doc.Footer.Default!.AddTable(1, 2);
                footerTable.Rows[0].Cells[0].Paragraphs[0].SetText("Owner");
                footerTable.Rows[0].Cells[1].Paragraphs[0].SetText("Security");
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding paragraph = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("Classification: Internal", paragraph.SourceText);
            Assert.Equal("Classification: Confidential", paragraph.TargetText);

            WordComparisonFinding cell = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("Platform", cell.SourceText);
            Assert.Equal("Security", cell.TargetText);
        }

        [Fact]
        public void CompareStructureReportsExternalLinkedImageChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_external_image.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph().AddImage(new Uri("https://example.com/logo-a.png"), 50, 50);
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_external_image.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph().AddImage(new Uri("https://example.com/logo-b.png"), 50, 50);
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding image = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("[Image: https://example.com/logo-a.png]", image.SourceText);
            Assert.Equal("[Image: https://example.com/logo-b.png]", image.TargetText);
        }

        [Fact]
        public void CompareStructureIncludesHeaderImages() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_header_image.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddHeadersAndFooters();
                doc.Header.Default!.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_header_image.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddHeadersAndFooters();
                doc.Header.Default!.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"));
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding image = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("image[0]", image.Location);
        }

        private static void AssertNoTempArtifact(WordDocument document) {
            Assert.Equal(string.Empty, document.FilePath);
        }
    }
}
