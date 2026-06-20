using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using OfficeIMO.Word;
using V = DocumentFormat.OpenXml.Vml;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CompareStructureAlignsBalancedParagraphInsertDeleteRanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_balanced_paragraph_range.docx");
            CreateDocumentWithParagraphsForReviewWave(sourcePath, "Terms", "Obsolete", "Closing");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_balanced_paragraph_range.docx");
            CreateDocumentWithParagraphsForReviewWave(targetPath, "Cover", "Terms updated", "Closing");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Cover");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Terms" &&
                finding.TargetText == "Terms updated");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted &&
                finding.SourceText == "Obsolete");
        }

        [Fact]
        public void CompareStructureDoesNotMoveImageWhenPrecedingXmlChangesOnly() {
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_image_position_preceding_xml.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Intro");
                doc.AddParagraph().AddImage(imagePath, 80, 40);
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_image_position_preceding_xml.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Intro");
                doc.AddParagraph().AddImage(imagePath, 80, 40);
                doc.Save(false);
            }

            ReplaceFirstBodyParagraphWithRuns(targetPath, "In", "tro");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.DoesNotContain(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.Message == "Image position changed.");
        }

        [Fact]
        public void CompareStructurePreservesTextOnlyHeaderParagraphStructure() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_text_header_structure.docx");
            CreateDocumentWithDefaultHeaderForReviewWave(sourcePath, "Shared header");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_text_header_structure.docx");
            CreateDocumentWithDefaultHeaderForReviewWave(targetPath, "Shared header");
            AddDuplicateDefaultHeaderWithParagraphs(targetPath, "Shared header", string.Empty);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == string.Empty);
        }

        [Fact]
        public void CompareStructureOrdersHeaderSnapshotsByReferences() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_header_reference_order.docx");
            CreateDocumentWithDefaultAndFirstHeaders(sourcePath, createFirstHeaderPartFirst: false);

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_header_reference_order.docx");
            CreateDocumentWithDefaultAndFirstHeaders(targetPath, createFirstHeaderPartFirst: true);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Empty(result.Findings);
        }

        [Fact]
        public void CompareStructureIgnoresVolatileVmlShapeIds() {
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_vml_shape_id.docx");
            CreateDocumentWithVmlImage(sourcePath, imagePath, "_x0000_s1025");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_vml_shape_id.docx");
            CreateDocumentWithVmlImage(targetPath, imagePath, "_x0000_s4096");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.DoesNotContain(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
        }

        [Fact]
        public void CompareStructureIgnoresVolatileDrawingEditIds() {
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_drawing_edit_id.docx");
            CreateDocumentWithDrawingEditIds(sourcePath, imagePath, "11111111", "22222222");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_drawing_edit_id.docx");
            CreateDocumentWithDrawingEditIds(targetPath, imagePath, "AAAAAAAA", "BBBBBBBB");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.DoesNotContain(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
        }

        [Fact]
        public void CompareStructureReportsNestedTableCellBlockReordering() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_cell_block_order.docx");
            CreateDocumentWithCellParagraphAndNestedTable(sourcePath, nestedTableFirst: false);

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_cell_block_order.docx");
            CreateDocumentWithCellParagraphAndNestedTable(targetPath, nestedTableFirst: true);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.Message == "Document block order changed.");
        }

        [Fact]
        public void CompareStructureReportsOmittedMergeValueShapeChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_omitted_merge.docx");
            CreateDocumentWithOneCellTableForReviewWave(sourcePath, "Merged");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_omitted_merge.docx");
            CreateDocumentWithOneCellTableForReviewWave(targetPath, "Merged");
            SetFirstCellOmittedHorizontalMerge(targetPath);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Merged" &&
                finding.TargetText == "Merged");
        }

        [Fact]
        public void CompareStructureAlignsBalancedTableRowInsertDeleteRanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_balanced_table_row_range.docx");
            CreateDocumentWithOneColumnTableForReviewWave(sourcePath, "Terms", "Obsolete", "Closing");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_balanced_table_row_range.docx");
            CreateDocumentWithOneColumnTableForReviewWave(targetPath, "Cover", "Terms updated", "Closing");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableRow &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Cover");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Terms" &&
                finding.TargetText == "Terms updated");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableRow &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted &&
                finding.SourceText == "Obsolete");
        }

        [Fact]
        public void CompareStructureReportsInlineImagePositionChangesInsideParagraphText() {
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_inline_image_offset.docx");
            CreateInlineImageParagraphDocument(sourcePath, imagePath, imageFirst: false);

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_inline_image_offset.docx");
            CreateInlineImageParagraphDocument(targetPath, imagePath, imageFirst: true);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.Message == "Image position changed.");
        }

        [Fact]
        public void CompareStructureReportsDrawingImageHyperlinkClickTargetChanges() {
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_drawing_image_hlink.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph().AddImage(imagePath, 80, 40);
                doc.Save(false);
            }

            AddDrawingImageHyperlinkClick(sourcePath, "https://evotec.xyz/source-click");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_drawing_image_hlink.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph().AddImage(imagePath, 80, 40);
                doc.Save(false);
            }

            AddDrawingImageHyperlinkClick(targetPath, "https://evotec.xyz/target-click");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
        }

        [Fact]
        public void CompareStructureIgnoresUnreferencedNormalNotes() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_unreferenced_note.docx");
            CreateDocumentWithParagraphsForReviewWave(sourcePath, "Policy");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_unreferenced_note.docx");
            CreateDocumentWithParagraphsForReviewWave(targetPath, "Policy");
            AddUnreferencedFootnote(targetPath, "Orphaned note text");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Empty(result.Findings);
        }

        [Fact]
        public void CompareStructureReportsParagraphNumberingStructureChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_numbering_structure.docx");
            CreateDocumentWithParagraphsForReviewWave(sourcePath, "Item");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_numbering_structure.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                WordList list = doc.AddList(WordListStyle.Numbered);
                list.AddItem("Item");
                doc.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Item" &&
                finding.TargetText == "Item");
        }

        [Fact]
        public void CompareStructureKeepsReferencedFootnotesInReferenceOrder() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_reference_order_notes.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Policy A").AddFootNote("Alpha note");
                doc.AddParagraph("Policy B").AddFootNote("Beta note");
                doc.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_reference_order_notes.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Policy A").AddFootNote("Alpha note");
                doc.AddParagraph("Policy B").AddFootNote("Beta note");
                doc.Save(false);
            }

            ReverseNormalFootnotePartOrder(targetPath);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Empty(result.Findings);
        }

        [Fact]
        public void CompareStructureReportsNumberingInstanceOverrideChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_numbering_override.docx");
            CreateDocumentWithNumberedItem(sourcePath, "Item");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_numbering_override.docx");
            CreateDocumentWithNumberedItem(targetPath, "Item");
            SetFirstNumberingInstanceStartOverride(targetPath, 0, 5);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Item" &&
                finding.TargetText == "Item");
        }

        [Fact]
        public void CompareStructureIgnoresInactiveAlternateContentFallbackImages() {
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_alternate_content_fallback.docx");
            CreateAlternateContentImageDocument(sourcePath, imagePath, "width:80pt;height:40pt");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_alternate_content_fallback.docx");
            CreateAlternateContentImageDocument(targetPath, imagePath, "width:120pt;height:40pt");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.DoesNotContain(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
        }

        [Fact]
        public void CompareStructureReportsTrackedDeletedTextInParagraphs() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_deleted_text.docx");
            CreateDocumentWithParagraphsForReviewWave(sourcePath, "Keep");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_deleted_text.docx");
            CreateDocumentWithDeletedTextParagraph(targetPath, "Keep", "Removed");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Keep" &&
                finding.TargetText == "Keep[Deleted:Removed]");
        }

        private static void CreateDocumentWithParagraphsForReviewWave(string path, params string[] paragraphs) {
            using WordDocument doc = WordDocument.Create(path);
            foreach (string paragraph in paragraphs) {
                doc.AddParagraph(paragraph);
            }

            doc.Save(false);
        }

        private static void CreateDocumentWithNumberedItem(string path, string text) {
            using WordDocument doc = WordDocument.Create(path);
            WordList list = doc.AddList(WordListStyle.Numbered);
            list.AddItem(text);
            doc.Save(false);
        }

        private static void CreateDocumentWithOneColumnTableForReviewWave(string path, params string[] rows) {
            using WordDocument doc = WordDocument.Create(path);
            WordTable table = doc.AddTable(rows.Length, 1);
            for (int row = 0; row < rows.Length; row++) {
                table.Rows[row].Cells[0].Paragraphs[0].SetText(rows[row]);
            }

            doc.Save(false);
        }

        private static void CreateDocumentWithDefaultHeaderForReviewWave(string path, string text) {
            using WordDocument doc = WordDocument.Create(path);
            doc.AddParagraph("Body");
            doc.HeaderDefaultOrCreate.AddParagraph(text);
            doc.Save(false);
        }

        private static void CreateDocumentWithOneCellTableForReviewWave(string path, string text) {
            using WordDocument doc = WordDocument.Create(path);
            WordTable table = doc.AddTable(1, 1);
            table.Rows[0].Cells[0].Paragraphs[0].SetText(text);
            doc.Save(false);
        }

        private static void ReplaceFirstBodyParagraphWithRuns(string path, string first, string second) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Paragraph paragraph = document.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
            paragraph.RemoveAllChildren<Run>();
            paragraph.Append(new Run(new Text(first)), new Run(new RunProperties(new Bold()), new Text(second)));
            document.MainDocumentPart.Document.Save();
        }

        private static void AddDuplicateDefaultHeaderWithParagraphs(string path, params string[] paragraphs) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = document.MainDocumentPart!;
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(paragraphs.Select(text => new Paragraph(new Run(new Text(text)))));
            headerPart.Header.Save();

            string relationshipId = mainPart.GetIdOfPart(headerPart);
            SectionProperties sectionProperties = GetOrCreateSectionProperties(mainPart);
            sectionProperties.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = relationshipId });
            mainPart.Document.Save();
        }

        private static void CreateDocumentWithDefaultAndFirstHeaders(string path, bool createFirstHeaderPartFirst) {
            using WordDocument doc = WordDocument.Create(path);
            doc.AddParagraph("Body");
            doc.Save(false);

            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = document.MainDocumentPart!;
            HeaderPart? defaultPart = null;
            HeaderPart? firstPart = null;

            if (createFirstHeaderPartFirst) {
                firstPart = AddHeaderPart(mainPart, "First header");
                defaultPart = AddHeaderPart(mainPart, "Default header");
            } else {
                defaultPart = AddHeaderPart(mainPart, "Default header");
                firstPart = AddHeaderPart(mainPart, "First header");
            }

            SectionProperties sectionProperties = GetOrCreateSectionProperties(mainPart);
            sectionProperties.RemoveAllChildren<HeaderReference>();
            sectionProperties.Append(new TitlePage());
            sectionProperties.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(defaultPart) });
            sectionProperties.Append(new HeaderReference { Type = HeaderFooterValues.First, Id = mainPart.GetIdOfPart(firstPart) });
            mainPart.Document.Save();
        }

        private static HeaderPart AddHeaderPart(MainDocumentPart mainPart, string text) {
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new Paragraph(new Run(new Text(text))));
            headerPart.Header.Save();
            return headerPart;
        }

        private static void CreateDocumentWithVmlImage(string path, string imagePath, string shapeId) {
            using WordDocument doc = WordDocument.Create(path);
            doc.AddParagraph("Body");
            doc.Save(false);

            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = document.MainDocumentPart!;
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (FileStream stream = File.OpenRead(imagePath)) {
                imagePart.FeedData(stream);
            }

            string relationshipId = mainPart.GetIdOfPart(imagePart);
            Body body = mainPart.Document.Body!;
            body.RemoveAllChildren<Paragraph>();
            body.PrependChild(new Paragraph(new Run(new Picture(
                new V.Shape(
                    new V.ImageData { RelationshipId = relationshipId }) {
                        Id = shapeId,
                        Style = "width:80pt;height:40pt",
                        Type = "#_x0000_t75"
                    }))));
            mainPart.Document.Save();
        }

        private static void CreateDocumentWithDrawingEditIds(string path, string imagePath, string editId, string anchorId) {
            using (WordDocument doc = WordDocument.Create(path)) {
                doc.AddParagraph().AddImage(imagePath, 80, 40);
                doc.Save(false);
            }

            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            DW.Inline inline = document.MainDocumentPart!.Document.Descendants<DW.Inline>().First();
            inline.SetAttribute(new OpenXmlAttribute("wp", "editId", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", editId));
            inline.SetAttribute(new OpenXmlAttribute("wp14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing", anchorId));
            document.MainDocumentPart.Document.Save();
        }

        private static void CreateDocumentWithCellParagraphAndNestedTable(string path, bool nestedTableFirst) {
            using WordDocument doc = WordDocument.Create(path);
            WordTable table = doc.AddTable(1, 1);
            table.Rows[0].Cells[0].Paragraphs[0].SetText("Cell paragraph");
            doc.Save(false);

            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            TableCell cell = document.MainDocumentPart!.Document.Descendants<TableCell>().First();
            Table nestedTable = new(
                new TableRow(
                    new TableCell(
                        new Paragraph(new Run(new Text("Nested value"))))));

            if (nestedTableFirst) {
                cell.PrependChild(nestedTable);
            } else {
                cell.Append(nestedTable);
            }

            document.MainDocumentPart.Document.Save();
        }

        private static void SetFirstCellOmittedHorizontalMerge(string path) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            TableCell cell = document.MainDocumentPart!.Document.Descendants<TableCell>().First();
            TableCellProperties properties = cell.GetFirstChild<TableCellProperties>() ?? cell.PrependChild(new TableCellProperties());
            properties.HorizontalMerge = new HorizontalMerge();
            document.MainDocumentPart.Document.Save();
        }

        private static void CreateInlineImageParagraphDocument(string path, string imagePath, bool imageFirst) {
            using (WordDocument doc = WordDocument.Create(path)) {
                doc.AddParagraph("Placeholder");
                doc.Save(false);
            }

            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = document.MainDocumentPart!;
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (FileStream stream = File.OpenRead(imagePath)) {
                imagePart.FeedData(stream);
            }

            string relationshipId = mainPart.GetIdOfPart(imagePart);
            Paragraph paragraph = imageFirst
                ? new Paragraph(
                    new Run(CreateInlineDrawing(relationshipId, 601U)),
                    new Run(new Text("Before After")))
                : new Paragraph(
                    new Run(new Text("Before ")),
                    new Run(CreateInlineDrawing(relationshipId, 601U)),
                    new Run(new Text(" After")));
            Body body = mainPart.Document.Body!;
            body.RemoveAllChildren<Paragraph>();
            body.PrependChild(paragraph);
            mainPart.Document.Save();
        }

        private static void AddDrawingImageHyperlinkClick(string path, string url) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = document.MainDocumentPart!;
            string relationshipId = mainPart.AddHyperlinkRelationship(new Uri(url), true).Id;
            PIC.NonVisualDrawingProperties properties = mainPart.Document.Descendants<PIC.NonVisualDrawingProperties>().First();
            properties.RemoveAllChildren<A.HyperlinkOnClick>();
            properties.Append(new A.HyperlinkOnClick { Id = relationshipId });
            mainPart.Document.Save();
        }

        private static void AddUnreferencedFootnote(string path, string text) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = document.MainDocumentPart!;
            FootnotesPart footnotesPart = mainPart.FootnotesPart ?? mainPart.AddNewPart<FootnotesPart>();
            footnotesPart.Footnotes ??= new Footnotes();
            long nextId = footnotesPart.Footnotes.Elements<Footnote>()
                .Where(footnote => footnote.Id?.Value != null)
                .Select(footnote => footnote.Id!.Value)
                .DefaultIfEmpty(1)
                .Max() + 1;
            footnotesPart.Footnotes.Append(new Footnote(new Paragraph(new Run(new Text(text)))) { Id = nextId });
            footnotesPart.Footnotes.Save();
        }

        private static void ReverseNormalFootnotePartOrder(string path) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Footnotes footnotes = document.MainDocumentPart!.FootnotesPart!.Footnotes!;
            List<Footnote> normalFootnotes = footnotes.Elements<Footnote>()
                .Where(footnote => footnote.Type == null || footnote.Type.Value == FootnoteEndnoteValues.Normal)
                .ToList();
            foreach (Footnote footnote in normalFootnotes) {
                footnote.Remove();
            }

            foreach (Footnote footnote in normalFootnotes.AsEnumerable().Reverse()) {
                footnotes.Append(footnote);
            }

            footnotes.Save();
        }

        private static void SetFirstNumberingInstanceStartOverride(string path, int level, int start) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Numbering numbering = document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
            NumberingInstance instance = numbering.Elements<NumberingInstance>().First();
            foreach (LevelOverride existingOverride in instance.Elements<LevelOverride>().Where(item => item.LevelIndex?.Value == level).ToList()) {
                existingOverride.Remove();
            }

            instance.Append(new LevelOverride(new StartOverrideNumberingValue { Val = start }) { LevelIndex = level });
            numbering.Save();
        }

        private static void CreateAlternateContentImageDocument(string path, string imagePath, string fallbackStyle) {
            using (WordDocument doc = WordDocument.Create(path)) {
                doc.AddParagraph("Placeholder");
                doc.Save(false);
            }

            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = document.MainDocumentPart!;
            string choiceRelationshipId = AddImagePart(mainPart, imagePath);
            string fallbackRelationshipId = AddImagePart(mainPart, imagePath);
            var alternateContent = new AlternateContent(
                new AlternateContentChoice(
                    new Run(CreateInlineDrawing(choiceRelationshipId, 701U))) { Requires = "wps" },
                new AlternateContentFallback(
                    new Run(new Picture(
                        new V.Shape(
                            new V.ImageData { RelationshipId = fallbackRelationshipId }) {
                                Id = "_x0000_s701",
                                Style = fallbackStyle,
                                Type = "#_x0000_t75"
                            }))));

            Body body = mainPart.Document.Body!;
            body.RemoveAllChildren<Paragraph>();
            body.PrependChild(new Paragraph(alternateContent));
            mainPart.Document.Save();
        }

        private static string AddImagePart(MainDocumentPart mainPart, string imagePath) {
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (FileStream stream = File.OpenRead(imagePath)) {
                imagePart.FeedData(stream);
            }

            return mainPart.GetIdOfPart(imagePart);
        }

        private static void CreateDocumentWithDeletedTextParagraph(string path, string text, string deletedText) {
            using WordDocument doc = WordDocument.Create(path);
            doc.AddParagraph("Placeholder");
            doc.Save(false);

            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Body body = document.MainDocumentPart!.Document.Body!;
            body.RemoveAllChildren<Paragraph>();
            body.PrependChild(new Paragraph(
                new Run(new Text(text)),
                new DeletedRun(new Run(new DeletedText(deletedText) { Space = SpaceProcessingModeValues.Preserve })) {
                    Author = "Codex",
                    Date = new DateTime(2026, 1, 1),
                    Id = "1"
                }));
            document.MainDocumentPart.Document.Save();
        }

        private static SectionProperties GetOrCreateSectionProperties(MainDocumentPart mainPart) {
            Body body = mainPart.Document.Body!;
            SectionProperties? sectionProperties = body.Elements<SectionProperties>().LastOrDefault();
            if (sectionProperties == null) {
                sectionProperties = new SectionProperties();
                body.Append(sectionProperties);
            }

            return sectionProperties;
        }
    }
}
