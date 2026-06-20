using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
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

        private static void CreateDocumentWithParagraphsForReviewWave(string path, params string[] paragraphs) {
            using WordDocument doc = WordDocument.Create(path);
            foreach (string paragraph in paragraphs) {
                doc.AddParagraph(paragraph);
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
