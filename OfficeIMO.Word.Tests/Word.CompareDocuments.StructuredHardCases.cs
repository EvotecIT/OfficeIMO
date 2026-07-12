using System;
using System.Collections.Generic;
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
        public void CompareStructureAlignsTableCellsUsingStructuralMatchText() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_table_literal_page_break_marker.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Before[PageBreak]After");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_table_real_page_break_marker.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("placeholder");
                doc.Save();
            }

            ReplaceFirstTableCellParagraph(targetPath, new Paragraph(
                new Run(new Text("Before")),
                new Run(new Break { Type = BreakValues.Page }),
                new Run(new Text("After"))));

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Before[PageBreak]After" &&
                finding.TargetText == "Before[PageBreak]After");
        }

        [Fact]
        public void CompareStructureReportsHyperlinkTargetChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_hyperlink_target.docx");
            CreateDocumentWithBodyText(sourcePath, "placeholder");
            ReplaceBodyWithHyperlink(sourcePath, "Portal", "https://evotec.xyz/source");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_hyperlink_target.docx");
            CreateDocumentWithBodyText(targetPath, "placeholder");
            ReplaceBodyWithHyperlink(targetPath, "Portal", "https://evotec.xyz/target");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Portal" &&
                finding.TargetText == "Portal");
        }

        [Fact]
        public void CompareStructureReportsTableCellHyperlinkTargetChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_cell_hyperlink_target.docx");
            CreateDocumentWithOneCellTable(sourcePath, "placeholder");
            ReplaceFirstTableCellWithHyperlink(sourcePath, "Portal", "https://evotec.xyz/source-cell");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_cell_hyperlink_target.docx");
            CreateDocumentWithOneCellTable(targetPath, "placeholder");
            ReplaceFirstTableCellWithHyperlink(targetPath, "Portal", "https://evotec.xyz/target-cell");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Portal" &&
                finding.TargetText == "Portal");
        }

        [Fact]
        public void CompareStructureAlignsTableRowsAfterMultipleInsertedRows() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_multi_inserted_rows.docx");
            CreateDocumentWithTableRows(sourcePath, "Owner: Alice", "Closing");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_multi_inserted_rows.docx");
            CreateDocumentWithTableRows(targetPath, "Evidence", "Summary", "Owner: Bob", "Closing");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableRow &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Evidence");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableRow &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Summary");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Owner: Alice" &&
                finding.TargetText == "Owner: Bob");
        }

        [Fact]
        public void CompareStructureReportsFieldInstructionChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_field_instruction.docx");
            CreateDocumentWithBodyText(sourcePath, "placeholder");
            ReplaceBodyParagraphForHardCase(sourcePath, CreateSimpleFieldParagraph("Project Alpha", " MERGEFIELD ProjectName "));

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_field_instruction.docx");
            CreateDocumentWithBodyText(targetPath, "placeholder");
            ReplaceBodyParagraphForHardCase(targetPath, CreateSimpleFieldParagraph("Project Alpha", " MERGEFIELD ProjectCode "));

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Project Alpha" &&
                finding.TargetText == "Project Alpha");
        }

        [Fact]
        public void CompareStructureIgnoresPackageLocalFootnoteIds() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_renumbered_footnote_id.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Policy").AddFootNote("Same note");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_renumbered_footnote_id.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Policy").AddFootNote("Same note");
                doc.Save();
            }

            RenumberFirstFootnote(targetPath, 42);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Empty(result.Findings);
        }

        [Fact]
        public void CompareStructureReportsFootnoteReferenceTargetChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_swapped_footnote_targets.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Policy A").AddFootNote("Alpha note");
                doc.AddParagraph("Policy B").AddFootNote("Beta note");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_swapped_footnote_targets.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Policy A").AddFootNote("Alpha note");
                doc.AddParagraph("Policy B").AddFootNote("Beta note");
                doc.Save();
            }

            SwapBodyFootnoteReferences(targetPath);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                (finding.SourceText?.Contains("[FootnoteReference:", StringComparison.Ordinal) ?? false) &&
                (finding.TargetText?.Contains("[FootnoteReference:", StringComparison.Ordinal) ?? false));
        }

        [Fact]
        public void CompareStructureReportsImageHyperlinkTargetChanges() {
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_image_hyperlink_target.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph().AddImage(imagePath, 80, 40);
                doc.Save();
            }

            WrapFirstDrawingInHyperlink(sourcePath, "https://evotec.xyz/source-image");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_image_hyperlink_target.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph().AddImage(imagePath, 80, 40);
                doc.Save();
            }

            WrapFirstDrawingInHyperlink(targetPath, "https://evotec.xyz/target-image");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
        }

        [Fact]
        public void CompareStructureReportsImageMovedWithinSamePart() {
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_image_moved_same_part.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph().AddImage(imagePath, 80, 40);
                doc.AddParagraph("Anchor text");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_image_moved_same_part.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Anchor text");
                doc.AddParagraph().AddImage(imagePath, 80, 40);
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.Message == "Image position changed.");
        }

        [Fact]
        public void CompareStructureReportsMergedCellShapeChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_merged_cell_shape.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                WordTable table = doc.AddTable(1, 2);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Project");
                table.Rows[0].Cells[1].Paragraphs[0].SetText("Status");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_merged_cell_shape.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 2);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Project");
                table.Rows[0].Cells[1].Paragraphs[0].SetText("Status");
                doc.Save();
            }

            SetFirstCellGridSpan(targetPath, 2);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Project" &&
                finding.TargetText == "Project");
        }

        [Fact]
        public void CompareStructureSkipsDuplicateEffectiveHeaderContent() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_duplicate_effective_header.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Body");
                doc.HeaderDefaultOrCreate.AddParagraph("Shared header");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_duplicate_effective_header.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Body");
                doc.HeaderDefaultOrCreate.AddParagraph("Shared header");
                doc.Save();
            }

            AddDuplicateDefaultHeaderReference(targetPath, "Shared header");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.DoesNotContain(result.Findings, finding =>
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                (finding.TargetText?.Contains("Shared header", StringComparison.Ordinal) ?? false));
        }

        [Fact]
        public void CompareStructureReportsDuplicateHeaderImageContentChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_duplicate_header_image.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Body");
                doc.HeaderDefaultOrCreate.AddParagraph("Shared header");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_duplicate_header_image.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Body");
                doc.HeaderDefaultOrCreate.AddParagraph("Shared header");
                doc.Save();
            }

            AddDuplicateDefaultHeaderReferenceWithImage(sourcePath, "Shared header", Path.Combine(_directoryWithImages, "EvotecLogo.png"));
            AddDuplicateDefaultHeaderReferenceWithImage(targetPath, "Shared header", Path.Combine(_directoryWithImages, "BackgroundImage.png"));

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
        }

        [Fact]
        public void CompareStructureReportsCrossScopeBlockReordering() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_cross_scope_order.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Approval");
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Terms");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_cross_scope_order.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordTable table = doc.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Terms");
                doc.AddParagraph("Approval");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.Message == "Document block order changed.");
        }

        [Fact]
        public void CompareStructureSkipsDisabledFirstPageHeaders() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_disabled_first_header.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Body");
                doc.HeaderFirstOrCreate.AddParagraph("Dormant first header");
                doc.Save();
            }

            RemoveTitlePageSwitch(sourcePath);

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_disabled_first_header.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Body");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.DoesNotContain(result.Findings, finding =>
                (finding.SourceText?.Contains("Dormant first header", StringComparison.Ordinal) ?? false) ||
                (finding.TargetText?.Contains("Dormant first header", StringComparison.Ordinal) ?? false));
        }

        [Fact]
        public void WordImageAssignsDrawingIdsAfterNoteImages() {
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string path = Path.Combine(_directoryWithFiles, "compare_structure_note_image_docpr_id.docx");
            using (WordDocument doc = WordDocument.Create(path)) {
                doc.AddParagraph("Policy").AddFootNote("Image note");
                doc.Save();
            }

            AppendFootnoteImageWithDrawingId(path, imagePath, 900U);

            using (WordDocument doc = WordDocument.Load(path)) {
                doc.AddParagraph().AddImage(imagePath, 80, 40);
                doc.Save();
            }

            List<uint> ids = GetDrawingDocPropertiesIds(path);

            Assert.Contains(900U, ids);
            Assert.Contains(ids, id => id > 900U);
            Assert.Equal(ids.Count, ids.Distinct().Count());
        }

        [Fact]
        public void WordImageAssignsDrawingIdsAfterCommentImages() {
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string path = Path.Combine(_directoryWithFiles, "compare_structure_comment_image_docpr_id.docx");
            using (WordDocument doc = WordDocument.Create(path)) {
                doc.AddParagraph("Policy");
                doc.Save();
            }

            AppendCommentImageWithDrawingId(path, imagePath, 900U);

            using (WordDocument doc = WordDocument.Load(path)) {
                doc.AddParagraph().AddImage(imagePath, 80, 40);
                doc.Save();
            }

            List<uint> ids = GetDrawingDocPropertiesIds(path);

            Assert.Contains(900U, ids);
            Assert.Contains(ids, id => id > 900U);
            Assert.Equal(ids.Count, ids.Distinct().Count());
        }

        private static void CreateDocumentWithBodyText(string path, string text) {
            using WordDocument doc = WordDocument.Create(path);
            doc.AddParagraph(text);
            doc.Save();
        }

        private static void CreateDocumentWithOneCellTable(string path, string text) {
            using WordDocument doc = WordDocument.Create(path);
            WordTable table = doc.AddTable(1, 1);
            table.Rows[0].Cells[0].Paragraphs[0].SetText(text);
            doc.Save();
        }

        private static void CreateDocumentWithTableRows(string path, params string[] rows) {
            using WordDocument doc = WordDocument.Create(path);
            WordTable table = doc.AddTable(rows.Length, 1);
            for (int index = 0; index < rows.Length; index++) {
                table.Rows[index].Cells[0].Paragraphs[0].SetText(rows[index]);
            }

            doc.Save();
        }

        private static void ReplaceBodyParagraphForHardCase(string path, Paragraph paragraph) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Body body = document.MainDocumentPart!.Document.Body!;
            body.RemoveAllChildren();
            body.Append(paragraph);
            document.MainDocumentPart.Document.Save();
        }

        private static void ReplaceBodyWithHyperlink(string path, string text, string url) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = document.MainDocumentPart!;
            string relationshipId = mainPart.AddHyperlinkRelationship(new Uri(url), true).Id;
            Body body = mainPart.Document.Body!;
            body.RemoveAllChildren();
            body.Append(new Paragraph(new Hyperlink(new Run(new Text(text))) { Id = relationshipId }));
            mainPart.Document.Save();
        }

        private static void ReplaceFirstTableCellWithHyperlink(string path, string text, string url) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = document.MainDocumentPart!;
            string relationshipId = mainPart.AddHyperlinkRelationship(new Uri(url), true).Id;
            TableCell cell = mainPart.Document.Descendants<TableCell>().First();
            cell.RemoveAllChildren<Paragraph>();
            cell.Append(new Paragraph(new Hyperlink(new Run(new Text(text))) { Id = relationshipId }));
            mainPart.Document.Save();
        }

        private static Paragraph CreateSimpleFieldParagraph(string displayText, string instruction) {
            return new Paragraph(new SimpleField(new Run(new Text(displayText))) { Instruction = instruction });
        }

        private static void RenumberFirstFootnote(string path, int newId) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Footnote footnote = document.MainDocumentPart!.FootnotesPart!.Footnotes!.Elements<Footnote>().First(item => item.Type == null);
            long oldId = footnote.Id!.Value;
            footnote.Id = newId;
            foreach (FootnoteReference reference in document.MainDocumentPart.Document.Descendants<FootnoteReference>().Where(item => item.Id?.Value == oldId)) {
                reference.Id = newId;
            }

            document.MainDocumentPart.FootnotesPart.Footnotes.Save();
            document.MainDocumentPart.Document.Save();
        }

        private static void SwapBodyFootnoteReferences(string path) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            List<FootnoteReference> references = document.MainDocumentPart!.Document.Descendants<FootnoteReference>().Take(2).ToList();
            long firstId = references[0].Id!.Value;
            references[0].Id = references[1].Id!.Value;
            references[1].Id = firstId;
            document.MainDocumentPart.Document.Save();
        }

        private static void WrapFirstDrawingInHyperlink(string path, string url) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = document.MainDocumentPart!;
            string relationshipId = mainPart.AddHyperlinkRelationship(new Uri(url), true).Id;
            DocumentFormat.OpenXml.Wordprocessing.Drawing drawing = mainPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().First();
            Run run = drawing.Ancestors<Run>().First();
            OpenXmlElement parent = run.Parent!;
            var hyperlink = new Hyperlink { Id = relationshipId };
            hyperlink.Append(run.CloneNode(true));
            parent.InsertBefore(hyperlink, run);
            run.Remove();
            mainPart.Document.Save();
        }

        private static void SetFirstCellGridSpan(string path, int span) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            TableCell cell = document.MainDocumentPart!.Document.Descendants<TableCell>().First();
            TableCellProperties properties = cell.GetFirstChild<TableCellProperties>() ?? cell.PrependChild(new TableCellProperties());
            properties.GridSpan = new GridSpan { Val = span };
            document.MainDocumentPart.Document.Save();
        }

        private static void AddDuplicateDefaultHeaderReference(string path, string text) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = document.MainDocumentPart!;
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new Paragraph(new Run(new Text(text))));
            headerPart.Header.Save();

            string relationshipId = mainPart.GetIdOfPart(headerPart);
            SectionProperties sectionProperties = mainPart.Document.Body!.Elements<SectionProperties>().Last();
            sectionProperties.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = relationshipId });
            mainPart.Document.Save();
        }

        private static void AddDuplicateDefaultHeaderReferenceWithImage(string path, string text, string imagePath) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = document.MainDocumentPart!;
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            ImagePart imagePart = Path.GetExtension(imagePath).Equals(".jpg", StringComparison.OrdinalIgnoreCase) ||
                                  Path.GetExtension(imagePath).Equals(".jpeg", StringComparison.OrdinalIgnoreCase)
                ? headerPart.AddImagePart(ImagePartType.Jpeg)
                : headerPart.AddImagePart(ImagePartType.Png);

            using (FileStream stream = File.OpenRead(imagePath)) {
                imagePart.FeedData(stream);
            }

            string imageRelationshipId = headerPart.GetIdOfPart(imagePart);
            headerPart.Header = new Header(
                new Paragraph(new Run(new Text(text))),
                new Paragraph(new Run(CreateInlineDrawing(imageRelationshipId, 77U))));
            headerPart.Header.Save();

            string relationshipId = mainPart.GetIdOfPart(headerPart);
            SectionProperties sectionProperties = mainPart.Document.Body!.Elements<SectionProperties>().Last();
            sectionProperties.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = relationshipId });
            mainPart.Document.Save();
        }

        private static void ReplaceFirstTableCellParagraph(string path, Paragraph paragraph) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            TableCell cell = document.MainDocumentPart!.Document.Descendants<TableCell>().First();
            cell.RemoveAllChildren<Paragraph>();
            cell.Append(paragraph);
            document.MainDocumentPart.Document.Save();
        }

        private static void RemoveTitlePageSwitch(string path) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            foreach (TitlePage titlePage in document.MainDocumentPart!.Document.Body!.Descendants<TitlePage>().ToList()) {
                titlePage.Remove();
            }

            document.MainDocumentPart.Document.Save();
        }

        private static void AppendFootnoteImageWithDrawingId(string path, string imagePath, uint drawingId) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            FootnotesPart footnotesPart = document.MainDocumentPart!.FootnotesPart!;
            Footnote footnote = footnotesPart.Footnotes!.Elements<Footnote>().First(item => item.Type == null);
            ImagePart imagePart = Path.GetExtension(imagePath).Equals(".jpg", StringComparison.OrdinalIgnoreCase) ||
                                  Path.GetExtension(imagePath).Equals(".jpeg", StringComparison.OrdinalIgnoreCase)
                ? footnotesPart.AddImagePart(ImagePartType.Jpeg)
                : footnotesPart.AddImagePart(ImagePartType.Png);

            using (FileStream stream = File.OpenRead(imagePath)) {
                imagePart.FeedData(stream);
            }

            string relationshipId = footnotesPart.GetIdOfPart(imagePart);
            footnote.Append(new Paragraph(new Run(CreateInlineDrawing(relationshipId, drawingId))));
            footnotesPart.Footnotes.Save();
        }

        private static void AppendCommentImageWithDrawingId(string path, string imagePath, uint drawingId) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = document.MainDocumentPart!;
            WordprocessingCommentsPart commentsPart = mainPart.WordprocessingCommentsPart ?? mainPart.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments ??= new Comments();
            ImagePart imagePart = Path.GetExtension(imagePath).Equals(".jpg", StringComparison.OrdinalIgnoreCase) ||
                                  Path.GetExtension(imagePath).Equals(".jpeg", StringComparison.OrdinalIgnoreCase)
                ? commentsPart.AddImagePart(ImagePartType.Jpeg)
                : commentsPart.AddImagePart(ImagePartType.Png);

            using (FileStream stream = File.OpenRead(imagePath)) {
                imagePart.FeedData(stream);
            }

            string relationshipId = commentsPart.GetIdOfPart(imagePart);
            commentsPart.Comments.Append(new Comment(new Paragraph(new Run(CreateInlineDrawing(relationshipId, drawingId)))) {
                Id = "0",
                Author = "OfficeIMO",
                Date = DateTime.UtcNow
            });
            commentsPart.Comments.Save();
        }

        private static DocumentFormat.OpenXml.Wordprocessing.Drawing CreateInlineDrawing(string relationshipId, uint id) {
            const long widthEmu = 914400L;
            const long heightEmu = 457200L;
            return new DocumentFormat.OpenXml.Wordprocessing.Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties { Id = id, Name = "Note image" },
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties { Id = id, Name = "Note image" },
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

        private static List<uint> GetDrawingDocPropertiesIds(string path) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, false);
            var roots = new OpenXmlElement?[] {
                document.MainDocumentPart!.RootElement,
                document.MainDocumentPart.FootnotesPart?.Footnotes,
                document.MainDocumentPart.EndnotesPart?.Endnotes,
                document.MainDocumentPart.WordprocessingCommentsPart?.Comments
            };

            return roots
                .Where(root => root != null)
                .SelectMany(root => root!.Descendants<DW.DocProperties>())
                .Where(properties => properties.Id != null)
                .Select(properties => properties.Id!.Value)
                .ToList();
        }
    }
}
