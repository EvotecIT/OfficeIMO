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
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_inserted_table_before_modified.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable summary = doc.AddTable(1, 1);
                summary.Rows[0].Cells[0].Paragraphs[0].SetText("Summary");
                WordTable terms = doc.AddTable(1, 1);
                terms.Rows[0].Cells[0].Paragraphs[0].SetText("Terms updated");
                WordTable closing = doc.AddTable(1, 1);
                closing.Rows[0].Cells[0].Paragraphs[0].SetText("Closing");
                doc.Save(false);
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
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_inserted_cell_before_modified.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 3);
                SetTableTexts(table, "Evidence", "Owner updated", "Closing");
                doc.Save(false);
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
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_inserted_image_before_modified.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph().AddImage(replacementPath, 120, 60);
                doc.AddParagraph().AddImage(replacementPath, 80, 40);
                doc.AddParagraph().AddImage(replacementPath, 50, 50);
                doc.Save(false);
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
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_first_header_variant.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.HeaderFirstOrCreate.AddParagraph("Classification: Confidential");
                doc.Save(false);
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
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_moved_footnote_marker.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Policy");
                doc.AddParagraph("Policy").AddFootNote("Policy note");
                doc.Save(false);
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
                doc.Save(false);
            }

            AppendTableAndImageToFirstFootnote(sourcePath, "Footnote table pending", logoPath);
            AppendTableAndImageToFirstEndnote(sourcePath, "Endnote table pending", logoPath);

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_note_tables_images.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Footnote anchor").AddFootNote("Footnote body");
                doc.AddParagraph("Endnote anchor").AddEndNote("Endnote body");
                doc.Save(false);
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
        public void CompareStructureSplitsChangedParagraphMovedAcrossParts() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_changed_body_to_header.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Status: Pending");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_changed_body_to_header.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.HeaderDefaultOrCreate.AddParagraph("Status: Approved");
                doc.Save(false);
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
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_picture_local_ids.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph().AddImage(imagePath, 80, 40);
                doc.Save(false);
            }

            MutatePictureLocalDrawingIds(targetPath, 4242U, "Different local picture id");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.DoesNotContain(result.Findings, finding => finding.Scope == WordComparisonScope.Image);
        }

        private static void AppendTableAndImageToFirstFootnote(string path, string tableText, string imagePath) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Footnote footnote = document.MainDocumentPart!.FootnotesPart!.Footnotes!.Elements<Footnote>().First(item => item.Type == null);
            footnote.Append(CreateOneCellTable(tableText));
            AppendNoteImage(document.MainDocumentPart.FootnotesPart, footnote, imagePath);
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
    }
}
