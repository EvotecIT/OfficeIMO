using System;
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
        public void CompareStructureMatchesFieldsAcrossInsertedEarlierField() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_field_insert_before_source.docx");
            CreateFieldRegressionDocument(sourcePath, (" AUTHOR ", "Alice"));

            string targetPath = Path.Combine(_directoryWithFiles, "compare_field_insert_before_target.docx");
            CreateFieldRegressionDocument(targetPath, (" TITLE ", "Plan"), (" AUTHOR ", "Alice"));

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                IncludedScopes = new System.Collections.Generic.HashSet<WordComparisonScope> { WordComparisonScope.Field }
            });

            WordComparisonFinding inserted = Assert.Single(result.Findings, finding => finding.Scope == WordComparisonScope.Field);
            Assert.Equal(WordComparisonChangeKind.Inserted, inserted.ChangeKind);
            Assert.Contains("TITLE", inserted.TargetText, StringComparison.Ordinal);
        }

        [Fact]
        public void CompareStructureMatchesContentControlsAcrossInsertedEarlierControl() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_control_insert_before_source.docx");
            CreateContentControlRegressionDocument(sourcePath, ("Stable", "Stable.Tag", "Stable value"));

            string targetPath = Path.Combine(_directoryWithFiles, "compare_control_insert_before_target.docx");
            CreateContentControlRegressionDocument(targetPath, ("New", "New.Tag", "Inserted value"), ("Stable", "Stable.Tag", "Stable value"));

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                IncludedScopes = new System.Collections.Generic.HashSet<WordComparisonScope> { WordComparisonScope.ContentControl }
            });

            WordComparisonFinding inserted = Assert.Single(result.Findings, finding => finding.Scope == WordComparisonScope.ContentControl);
            Assert.Equal(WordComparisonChangeKind.Inserted, inserted.ChangeKind);
            Assert.Contains("alias=New", inserted.TargetText, StringComparison.Ordinal);
        }

        [Fact]
        public void CompareStructureMatchesRevisionsAcrossInsertedEarlierRevision() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_revision_insert_before_source.docx");
            CreateRevisionRegressionDocument(sourcePath, "Keep");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_revision_insert_before_target.docx");
            CreateRevisionRegressionDocument(targetPath, "New", "Keep");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                CompareGeneratedIds = false,
                CompareVolatileMetadata = false,
                IncludedScopes = new System.Collections.Generic.HashSet<WordComparisonScope> { WordComparisonScope.Revision }
            });

            WordComparisonFinding inserted = Assert.Single(result.Findings, finding => finding.Scope == WordComparisonScope.Revision);
            Assert.Equal(WordComparisonChangeKind.Inserted, inserted.ChangeKind);
            Assert.Contains("text=New", inserted.TargetText, StringComparison.Ordinal);
        }

        [Fact]
        public void CompareStructureReportArtifactTracksFeatureFindingsWhenTextFindingsAreDisabled() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_report_artifact_feature_source.docx");
            CreateFieldRegressionDocument(sourcePath, (" AUTHOR ", "Alice"));

            string targetPath = Path.Combine(_directoryWithFiles, "compare_report_artifact_feature_target.docx");
            CreateFieldRegressionDocument(targetPath, (" TITLE ", "Plan"));

            string outputPath = Path.Combine(_directoryWithFiles, "compare_report_artifact_feature_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    TrackTextFindings = false,
                    TrackFeatureFindings = true
                });

            Assert.Contains(result.Findings, finding => finding.Scope == WordComparisonScope.Field);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            string redlineText = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!.InnerText;
            Assert.Contains("Tracked Changes", redlineText, StringComparison.Ordinal);
            Assert.Contains("Field", redlineText, StringComparison.Ordinal);
        }

        [Fact]
        public void CompareStructureInPlaceRunFormattingUsesMatchedSourceParagraphAfterInsertedParagraph() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_run_format_insert_before_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Stable run");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_run_format_insert_before_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Inserted before");
                document.AddParagraph("Stable run").Bold = true;
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_run_format_insert_before_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding => finding.Message == "Run formatting changed.");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Paragraph changedParagraph = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Last();
            Run changedRun = Assert.Single(changedParagraph.Elements<Run>());
            Assert.NotNull(changedRun.RunProperties?.RunPropertiesChange);
        }

        [Fact]
        public void CompareStructureInPlaceRedlineSkipsImageOnlyParagraphIndexes() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_image_only_paragraph_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"), 80, 40);
                document.AddParagraph("Stable text");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_image_only_paragraph_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"), 80, 40);
                document.AddParagraph("Changed text");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_image_only_paragraph_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Paragraph[] paragraphs = redline._document.Body!.Elements<Paragraph>().ToArray();
            Assert.True(paragraphs[0].Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any());
            Assert.Empty(paragraphs[0].Descendants<InsertedRun>());
            Assert.Empty(paragraphs[0].Descendants<DeletedRun>());
            Assert.Contains(paragraphs[1].Descendants<DeletedRun>(), run => run.InnerText == "Stable text");
            Assert.Contains(paragraphs[1].Descendants<InsertedRun>(), run => run.InnerText == "Changed text");
        }

        [Fact]
        public void CompareStructureInPlaceRedlinesParagraphFormattingChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_paragraph_format_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Formatted paragraph");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_paragraph_format_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Formatted paragraph").SetStyle(WordParagraphStyles.Heading1);
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_paragraph_format_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding => finding.Message == "Paragraph style id changed.");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Paragraph paragraph = Assert.Single(redline._document.Body!.Elements<Paragraph>(), item => item.InnerText == "Formatted paragraph");
            Assert.NotNull(paragraph.ParagraphProperties?.ParagraphPropertiesChange);
        }

        [Fact]
        public void CompareStructureInPlaceRunFormattingUsesOriginalTargetParagraphsAfterDeletedParagraph() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_run_format_after_deleted_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Removed before");
                document.AddParagraph("Stable run");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_run_format_after_deleted_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Stable run").Bold = true;
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_run_format_after_deleted_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding => finding.Message == "Paragraph deleted.");
            Assert.Contains(result.Findings, finding => finding.Message == "Run formatting changed.");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Paragraph stableParagraph = Assert.Single(redline._document.Body!.Elements<Paragraph>(), paragraph => paragraph.InnerText == "Stable run");
            Run stableRun = Assert.Single(stableParagraph.Elements<Run>());
            Assert.NotNull(stableRun.RunProperties?.RunPropertiesChange);
        }

        [Fact]
        public void CompareStructureInPlaceRunFormattingUsesDuplicateParagraphPositionBeforeTextFallback() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_duplicate_run_format_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Same");
                document.AddParagraph("Same").Italic = true;
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_duplicate_run_format_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Same");
                document.AddParagraph("Same").Bold = true;
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_duplicate_run_format_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Paragraph[] paragraphs = redline._document.Body!.Elements<Paragraph>().ToArray();
            Run firstRun = Assert.Single(paragraphs[0].Elements<Run>());
            Run secondRun = Assert.Single(paragraphs[1].Elements<Run>());
            Assert.Null(firstRun.RunProperties?.RunPropertiesChange);
            PreviousRunProperties previousProperties = Assert.IsType<PreviousRunProperties>(secondRun.RunProperties?.RunPropertiesChange?.FirstChild);
            Assert.NotNull(previousProperties.GetFirstChild<Italic>());
        }

        [Fact]
        public void CompareStructureInPlaceRunFormattingUsesOverlappingSourceRunForResegmentedText() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_resegmented_run_format_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document._document.Body!.Append(new Paragraph(
                    new Run(new Text("A") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(
                        new RunProperties(new Bold()),
                        new Text("B") { Space = SpaceProcessingModeValues.Preserve })));
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_resegmented_run_format_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document._document.Body!.Append(new Paragraph(
                    new Run(new Text("AB") { Space = SpaceProcessingModeValues.Preserve })));
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_resegmented_run_format_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding => finding.Message == "Run formatting changed.");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Paragraph changedParagraph = Assert.Single(redline._document.Body!.Elements<Paragraph>(), paragraph => paragraph.InnerText == "AB");
            Run changedRun = Assert.Single(changedParagraph.Elements<Run>());
            PreviousRunProperties previousProperties = Assert.IsType<PreviousRunProperties>(changedRun.RunProperties?.RunPropertiesChange?.FirstChild);
            Assert.NotNull(previousProperties.GetFirstChild<Bold>());
        }

        [Fact]
        public void CompareStructureInPlacePreservesConsecutiveDeletedImageOrder() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_deleted_image_order_source.docx");
            string firstImage = Path.Combine(_directoryWithImages, "Kulek.jpg");
            string secondImage = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string stableImage = Path.Combine(_directoryWithImages, "snail.bmp");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph().AddImage(firstImage, 40, 40);
                document.AddParagraph().AddImage(secondImage, 40, 40);
                document.AddParagraph().AddImage(stableImage, 40, 40);
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_deleted_image_order_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph().AddImage(stableImage, 40, 40);
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_deleted_image_order_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            MainDocumentPart mainPart = redline._wordprocessingDocument.MainDocumentPart!;
            List<byte[]> deletedImages = GetDeletedDrawingImageBytes(mainPart, mainPart.Document.Body!);
            Assert.Equal(2, deletedImages.Count);
            Assert.Equal(File.ReadAllBytes(firstImage), deletedImages[0]);
            Assert.Equal(File.ReadAllBytes(secondImage), deletedImages[1]);
        }

        [Fact]
        public void CompareStructureInPlaceInsertsDeletedImageBeforeFinalSectionProperties() {
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_deleted_image_sectpr_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph().AddImage(imagePath, 40, 40);
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_deleted_image_sectpr_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Target keeps only text");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_deleted_image_sectpr_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            using WordprocessingDocument package = WordprocessingDocument.Open(outputPath, false);
            Body body = package.MainDocumentPart!.Document.Body!;
            Assert.IsType<SectionProperties>(body.ChildElements.Last());
            Paragraph deletedImageParagraph = Assert.Single(body.Elements<Paragraph>(), paragraph => paragraph.Descendants<DeletedRun>().Any());
            Assert.True(body.ChildElements.ToList().IndexOf(deletedImageParagraph) < body.ChildElements.Count - 1);
        }

        [Fact]
        public void CompareStructureInPlaceMatchesDeletedImagesByPart() {
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_deleted_header_image_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddHeadersAndFooters();
                document.Header.Default!.AddParagraph().AddImage(imagePath, 40, 40);
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_deleted_header_image_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddHeadersAndFooters();
                document.Header.Default!.AddParagraph("Header remains");
                document.AddParagraph().AddImage(imagePath, 40, 40);
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_deleted_header_image_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            using WordprocessingDocument package = WordprocessingDocument.Open(outputPath, false);
            HeaderPart headerPart = Assert.Single(package.MainDocumentPart!.HeaderParts);
            Assert.NotEmpty(headerPart.Header.Descendants<DeletedRun>());
            Assert.Empty(package.MainDocumentPart.Document.Body!.Descendants<DeletedRun>());
        }

        [Fact]
        public void CompareStructureInPlaceInsertsDeletedParagraphWhenTargetPartHasNoParagraphEntries() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_deleted_header_empty_part_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddHeadersAndFooters();
                Header header = document._wordprocessingDocument.MainDocumentPart!.HeaderParts.Single().Header!;
                header.RemoveAllChildren();
                header.Append(new Paragraph(new Run(new Text("Deleted header paragraph") { Space = SpaceProcessingModeValues.Preserve })));
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_deleted_header_empty_part_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddHeadersAndFooters();
                Header header = document._wordprocessingDocument.MainDocumentPart!.HeaderParts.Single().Header!;
                header.RemoveAllChildren();
                header.Append(CreateSingleCellTable("Header table remains"));
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_deleted_header_empty_part_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            using WordprocessingDocument package = WordprocessingDocument.Open(outputPath, false);
            HeaderPart headerPart = Assert.Single(package.MainDocumentPart!.HeaderParts);
            Assert.Contains(headerPart.Header.Descendants<DeletedRun>(), run => run.InnerText == "Deleted header paragraph");
            Assert.DoesNotContain(package.MainDocumentPart.Document.Body!.Descendants<DeletedRun>(), run => run.InnerText == "Deleted header paragraph");
        }

        [Fact]
        public void CompareStructureInPlacePreservesConsecutiveDeletedTableOrder() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_deleted_table_order_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "Deleted A";
                document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "Deleted B";
                document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "Stable C";
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_deleted_table_order_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "Stable C";
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_deleted_table_order_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            string[] tableTexts = redline._document.Body!.Elements<Table>().Select(table => table.InnerText).ToArray();
            Assert.Equal(new[] { "Deleted A", "Deleted B", "Stable C" }, tableTexts);
        }

        [Fact]
        public void CompareStructureInPlaceMapsDeletedNestedTableToSurvivingParentTable() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_deleted_nested_table_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document._document.Body!.Append(CreateSingleCellTable("Deleted top table"));
                document._document.Body!.Append(CreateSingleCellTable(
                    "Parent stable",
                    CreateSingleCellTable("Nested deleted")));
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_deleted_nested_table_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document._document.Body!.Append(CreateSingleCellTable("Parent stable"));
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_deleted_nested_table_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Table[] topLevelTables = redline._document.Body!.Elements<Table>().ToArray();
            Assert.Equal(2, topLevelTables.Length);
            Assert.Contains(topLevelTables[0].Descendants<DeletedRun>(), run => run.InnerText == "Deleted top table");
            Table nestedTable = Assert.Single(topLevelTables[1].Descendants<Table>());
            Assert.Contains(nestedTable.Descendants<DeletedRun>(), run => run.InnerText == "Nested deleted");
        }

        [Fact]
        public void CompareStructureInPlaceAppliesTableCellRedlinesBeforeDeletedTableInsertion() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_deleted_table_before_cell_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "Deleted table";
                document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "Stable old";
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_deleted_table_before_cell_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "Stable new";
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_deleted_table_before_cell_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Table[] tables = redline._document.Body!.Elements<Table>().ToArray();
            Assert.Equal(2, tables.Length);
            Assert.Equal("Deleted table", tables[0].InnerText);
            Assert.Contains(tables[1].Descendants<DeletedRun>(), run => run.InnerText == "Stable old");
            Assert.Contains(tables[1].Descendants<InsertedRun>(), run => run.InnerText == "Stable new");
        }

        [Fact]
        public void CompareStructureInPlaceInsertsDeletedContentControlsIntoOriginalPart() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_deleted_header_sdt_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddHeadersAndFooters();
                document.Header.Default!.AddParagraph().AddStructuredDocumentTag("Header deleted", "Header", "HeaderTag");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_deleted_header_sdt_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddHeadersAndFooters();
                document.Header.Default!.AddParagraph("Header remains");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_deleted_header_sdt_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            using WordprocessingDocument package = WordprocessingDocument.Open(outputPath, false);
            HeaderPart headerPart = Assert.Single(package.MainDocumentPart!.HeaderParts);
            Assert.Contains(headerPart.Header.Descendants<DeletedRun>(), deleted => deleted.InnerText == "Header deleted");
            Assert.DoesNotContain(package.MainDocumentPart.Document.Body!.Descendants<DeletedRun>(), deleted => deleted.InnerText == "Header deleted");
        }

        [Fact]
        public void CompareStructureInPlaceInsertsDeletedContentControlsIntoOriginalFootnote() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_deleted_footnote_sdt_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Footnote anchor").AddFootNote("Footnote body");
                document.Save(false);
            }

            AddContentControlToFirstFootnote(sourcePath, "Deleted note control");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_deleted_footnote_sdt_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Footnote anchor").AddFootNote("Footnote body");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_deleted_footnote_sdt_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            using WordprocessingDocument package = WordprocessingDocument.Open(outputPath, false);
            Footnotes footnotes = package.MainDocumentPart!.FootnotesPart!.Footnotes!;
            Footnote note = footnotes.Elements<Footnote>().First(footnote => footnote.Type == null || footnote.Type.Value == FootnoteEndnoteValues.Normal);
            Assert.Contains(note.Descendants<DeletedRun>(), deleted => deleted.InnerText == "Deleted note control");
            Assert.Empty(footnotes.Elements<SdtBlock>());
            Assert.Empty(footnotes.Elements<Paragraph>());
        }

        [Fact]
        public void CompareStructureInPlaceCopiesDeletedImageHyperlinkRelationships() {
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string stableImagePath = Path.Combine(_directoryWithImages, "snail.bmp");
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_deleted_image_hyperlink_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph().AddImage(imagePath, 80, 40);
                document.AddParagraph().AddImage(stableImagePath, 40, 40);
                document.Save(false);
            }

            AddDrawingImageHyperlinkClick(sourcePath, "https://evotec.xyz/deleted-image");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_deleted_image_hyperlink_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph().AddImage(stableImagePath, 40, 40);
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_deleted_image_hyperlink_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            using WordprocessingDocument package = WordprocessingDocument.Open(outputPath, false);
            MainDocumentPart mainPart = package.MainDocumentPart!;
            DeletedRun deletedRun = Assert.Single(mainPart.Document.Body!.Descendants<DeletedRun>());
            DocumentFormat.OpenXml.Drawing.HyperlinkOnClick hyperlink = Assert.Single(deletedRun.Descendants<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>());
            Assert.NotNull(hyperlink.Id?.Value);
            string relationshipId = hyperlink.Id!.Value!;
            HyperlinkRelationship relationship = Assert.Single(mainPart.HyperlinkRelationships, item => item.Id == relationshipId);
            Assert.Equal(new Uri("https://evotec.xyz/deleted-image"), relationship.Uri);
        }

        [Fact]
        public void CompareStructureInPlaceParagraphRewriteRemovesTargetHyperlinkTextContainers() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_paragraph_hyperlink_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document._document.Body!.Append(new Paragraph(new Run(new Text("Original link text") { Space = SpaceProcessingModeValues.Preserve })));
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_paragraph_hyperlink_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document._document.Body!.Append(new Paragraph(
                    new Hyperlink(
                        new Run(new Text("Changed link text") { Space = SpaceProcessingModeValues.Preserve })) {
                        Anchor = "ChangedBookmark"
                    }));
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_paragraph_hyperlink_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new System.Collections.Generic.HashSet<WordComparisonScope> { WordComparisonScope.Paragraph }
                    }
                });

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Paragraph paragraph = Assert.Single(redline._document.Body!.Elements<Paragraph>());
            Assert.Empty(paragraph.Descendants<Hyperlink>());
            Assert.Equal(1, paragraph.Descendants<InsertedRun>().Count(run => run.InnerText == "Changed link text"));
        }

        [Fact]
        public void CompareStructureInPlacePlacesTrailingDeletedRunAfterLastTargetRun() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_trailing_deleted_run_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document._document.Body!.Append(new Paragraph(
                    new Run(new Text("Keep") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new Text(" deleted") { Space = SpaceProcessingModeValues.Preserve })));
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_trailing_deleted_run_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document._document.Body!.Append(new Paragraph(new Run(new Text("Keep") { Space = SpaceProcessingModeValues.Preserve })));
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_trailing_deleted_run_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new System.Collections.Generic.HashSet<WordComparisonScope> { WordComparisonScope.Run }
                    }
                });

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Paragraph paragraph = Assert.Single(redline._document.Body!.Elements<Paragraph>(), item => item.InnerText.Contains("Keep", StringComparison.Ordinal));
            OpenXmlElement[] textChildren = paragraph.ChildElements
                .Where(child => child is Run || child is DeletedRun)
                .ToArray();

            Assert.IsType<Run>(textChildren[0]);
            Assert.IsType<DeletedRun>(textChildren[1]);
            Assert.Equal("Keep", textChildren[0].InnerText);
            Assert.Equal(" deleted", textChildren[1].InnerText);
        }

        [Fact]
        public void CompareStructureInPlaceRedlinesOuterContentControlWhenNestedControlAlsoChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_nested_sdt_parent_source.docx");
            CreateNestedContentControlDocument(sourcePath, "Outer old", "Child old");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_nested_sdt_parent_target.docx");
            CreateNestedContentControlDocument(targetPath, "Outer new", "Child new");

            string outputPath = Path.Combine(_directoryWithFiles, "compare_nested_sdt_parent_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new System.Collections.Generic.HashSet<WordComparisonScope> { WordComparisonScope.ContentControl }
                    }
                });

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            SdtBlock control = Assert.Single(redline._document.Body!.Descendants<SdtBlock>());
            Assert.Contains(control.Descendants<DeletedRun>(), run => run.InnerText.Contains("Outer old", StringComparison.Ordinal));
            Assert.Contains(control.Descendants<DeletedRun>(), run => run.InnerText.Contains("Child old", StringComparison.Ordinal));
            Assert.Contains(control.Descendants<InsertedRun>(), run => run.InnerText.Contains("Outer new", StringComparison.Ordinal));
            Assert.Contains(control.Descendants<InsertedRun>(), run => run.InnerText.Contains("Child new", StringComparison.Ordinal));
        }

        [Fact]
        public void CompareStructureInPlaceSkipsNestedDeletedContentControlWhenParentIsDeleted() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_nested_sdt_deleted_source.docx");
            CreateNestedContentControlDocument(sourcePath, "Deleted outer", "Deleted child");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_nested_sdt_deleted_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("No content controls remain.");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_nested_sdt_deleted_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new System.Collections.Generic.HashSet<WordComparisonScope> { WordComparisonScope.ContentControl }
                    }
                });

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Assert.Equal(1, redline._document.Body!.Descendants<DeletedRun>().Count(run => run.InnerText.Contains("Deleted child", StringComparison.Ordinal)));
        }

        [Fact]
        public void CompareStructureKeysFootnoteParagraphsByNoteId() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_footnote_id_source.docx");
            CreateFootnoteIdRegressionDocument(
                sourcePath,
                (5, "Deleted anchor", "Deleted footnote"),
                (9, "Stable anchor", "Stable footnote"));

            string targetPath = Path.Combine(_directoryWithFiles, "compare_footnote_id_target.docx");
            CreateFootnoteIdRegressionDocument(targetPath, (9, "Stable anchor", "Stable footnote"));

            WordComparisonResult result = WordDocumentComparer.CompareStructure(
                sourcePath,
                targetPath,
                new WordComparisonOptions {
                    IncludedScopes = new System.Collections.Generic.HashSet<WordComparisonScope> { WordComparisonScope.Paragraph }
                });

            Assert.Contains(result.Findings, finding =>
                finding.ChangeKind == WordComparisonChangeKind.Deleted &&
                string.Equals(finding.SourceText, "Deleted footnote", StringComparison.Ordinal));
            Assert.DoesNotContain(result.Findings, finding =>
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                ((finding.SourceText?.Contains("Stable footnote", StringComparison.Ordinal) == true) ||
                 (finding.TargetText?.Contains("Stable footnote", StringComparison.Ordinal) == true)));
        }

        [Fact]
        public void TableOfContentRefreshCombinesSplitComplexTocInstruction() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshSplitComplexInstruction.docx");

            using WordDocument document = WordDocument.Create(filePath);
            WordTableOfContent toc = document.AddTableOfContent();
            ReplaceTocInstructionWithSplitComplexField(toc);
            document._document.Body!.Append(CreateTcFieldParagraph("Alpha", "A"));
            document._document.Body!.Append(CreateTcFieldParagraph("Beta", "B"));

            WordTableOfContentRefreshReport report = toc.RefreshEntries();

            Assert.Equal(1, report.EntryCount);
            Assert.Equal("Alpha", Assert.Single(report.Entries).Text);
            Assert.DoesNotContain("Beta", TocText(toc));
        }

        [Fact]
        public void TableOfContentRefreshIndexCombinesSplitComplexIndexInstruction() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshSplitComplexInstruction.docx");

            using WordDocument document = WordDocument.Create(filePath);
            WordTableOfContent index = document.AddTableOfContent();
            ReplaceIndexInstructionWithSplitComplexField(index);
            AddIndexEntryParagraph(document, "Alpha topic", " XE \"Alpha\" \\f \"A\" ");
            AddIndexEntryParagraph(document, "Beta topic", " XE \"Beta\" \\f \"B\" ");

            WordIndexRefreshReport report = index.RefreshIndex("Split Complex Index");

            WordIndexEntry entry = Assert.Single(report.Entries);
            Assert.Equal("Alpha", entry.Term);
            Assert.Equal(2, report.ColumnCount);
            Assert.Contains("Alpha", TocText(index));
            Assert.DoesNotContain("Beta", TocText(index));
        }

        [Fact]
        public void ImportedRawComplexTocStartsAtMatchedFieldWhenEarlierFieldExists() {
            string filePath = Path.Combine(_directoryWithFiles, "ImportedRawComplexTocAfterAuthor.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document._document.Body!.RemoveAllChildren<Paragraph>();
                document._document.Body!.Append(CreateParagraphWithAuthorAndRawTocFields());
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains(document._document.Body!.Elements<Paragraph>(), paragraph =>
                    paragraph.Descendants<FieldCode>().Any(fieldCode => fieldCode.Text.Contains("AUTHOR", StringComparison.Ordinal)));
                Assert.Contains(toc.SdtBlock.Descendants<FieldCode>(), fieldCode =>
                    fieldCode.Text.Contains("TOC", StringComparison.Ordinal));
                Assert.DoesNotContain(toc.SdtBlock.Descendants<FieldCode>(), fieldCode =>
                    fieldCode.Text.Contains("AUTHOR", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void ImportedRawSimpleTocLeavesLeadingTextOutsideGeneratedSdt() {
            string filePath = Path.Combine(_directoryWithFiles, "ImportedRawSimpleTocLeadingText.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document._document.Body!.RemoveAllChildren<Paragraph>();
                document._document.Body!.Append(new Paragraph(
                    new Run(new Text("Before raw TOC") { Space = SpaceProcessingModeValues.Preserve }),
                    new SimpleField(new Run(new Text("No entries") { Space = SpaceProcessingModeValues.Preserve })) {
                        Instruction = " TOC \\o \"1-3\" \\h \\z \\u "
                    }));
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                OpenXmlElement[] bodyChildren = document._document.Body!.ChildElements.ToArray();

                Assert.IsType<Paragraph>(bodyChildren[0]);
                Assert.Equal("Before raw TOC", bodyChildren[0].InnerText);
                Assert.IsType<SdtBlock>(bodyChildren[1]);
                Assert.DoesNotContain("Before raw TOC", toc.SdtBlock.InnerText, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void ImportedRawComplexTocLeavesSurroundingTextOutsideGeneratedSdt() {
            string filePath = Path.Combine(_directoryWithFiles, "ImportedRawComplexTocSurroundingText.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document._document.Body!.RemoveAllChildren<Paragraph>();
                document._document.Body!.Append(new Paragraph(
                    new Run(new Text("Before complex TOC") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" TOC \\o \"1-3\" \\h \\z \\u ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("No entries") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                    new Run(new Text("After complex TOC") { Space = SpaceProcessingModeValues.Preserve })));
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                OpenXmlElement[] bodyChildren = document._document.Body!.ChildElements.ToArray();

                Assert.IsType<Paragraph>(bodyChildren[0]);
                Assert.Equal("Before complex TOC", bodyChildren[0].InnerText);
                Assert.IsType<SdtBlock>(bodyChildren[1]);
                Assert.IsType<Paragraph>(bodyChildren[2]);
                Assert.Equal("After complex TOC", bodyChildren[2].InnerText);
                Assert.DoesNotContain("Before complex TOC", toc.SdtBlock.InnerText, StringComparison.Ordinal);
                Assert.DoesNotContain("After complex TOC", toc.SdtBlock.InnerText, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void CompareStructureInPlaceRedlinesReviewFindings() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_review_in_place_source.docx");
            CreateRevisionRegressionDocument(sourcePath);

            string targetPath = Path.Combine(_directoryWithFiles, "compare_review_in_place_target.docx");
            CreateRevisionRegressionDocument(targetPath, "New review text");

            string outputPath = Path.Combine(_directoryWithFiles, "compare_review_in_place_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "Redline Bot",
                    ComparisonOptions = new WordComparisonOptions {
                        CompareGeneratedIds = false,
                        CompareVolatileMetadata = false,
                        IncludedScopes = new System.Collections.Generic.HashSet<WordComparisonScope> { WordComparisonScope.Revision }
                    }
                });

            Assert.Contains(result.Findings, finding => finding.Scope == WordComparisonScope.Revision);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Assert.Contains("Tracked Review Changes", redline._document.Body!.InnerText, StringComparison.Ordinal);
            Assert.Contains(redline._document.Body!.Descendants<InsertedRun>(), run =>
                run.Author?.Value == "Redline Bot" &&
                run.InnerText.Contains("New review text", StringComparison.Ordinal));
        }

        [Fact]
        public void CompareStructureInPlaceAllowsSignedTargetCopies() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_signed_target_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Original text");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_signed_target_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Changed text");
                document.Save(false);
            }

            AddDigitalSignatureMetadata(targetPath, CreateSignatureXml());

            string outputPath = Path.Combine(_directoryWithFiles, "compare_signed_target_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Assert.True(redline.InspectSignatures().HasSignatures);
            Assert.Contains(redline._document.Body!.Descendants<InsertedRun>(), run => run.InnerText.Contains("Changed text", StringComparison.Ordinal));
        }

        [Fact]
        public void UpdateFieldsAndGetReportParsesNumericPictureBeforeTrailingFormatSwitches() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.NumericPictureBeforeMergeFormat.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document._document.Body!.Append(new Paragraph(
                    new SimpleField(new Run(new Text("stale") { Space = SpaceProcessingModeValues.Preserve })) {
                        Instruction = " PAGE \\# \"000\" \\* MERGEFORMAT "
                    }));
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                WordFieldUpdateResult update = Assert.Single(report.Results);
                Assert.Equal(WordFieldUpdateStatus.Updated, update.Status);
                Assert.Equal("001", update.ResultText);
                Assert.Equal(0, report.ParseErrorCount);
            }
        }

        [Fact]
        public void CompareStructureInPlaceRedlineKeepsConsecutiveDeletedParagraphOrder() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_deleted_paragraph_order_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("A");
                document.AddParagraph("B");
                document.AddParagraph("C");
                document.AddParagraph("D");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_deleted_paragraph_order_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("A");
                document.AddParagraph("D");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_deleted_paragraph_order_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            string[] paragraphTexts = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
                .Select(paragraph => paragraph.InnerText)
                .Where(text => text is "A" or "B" or "C" or "D")
                .ToArray();

            Assert.Equal(new[] { "A", "B", "C", "D" }, paragraphTexts);
            Assert.Contains(redline._document.Body!.Descendants<DeletedRun>(), run => run.InnerText == "B");
            Assert.Contains(redline._document.Body!.Descendants<DeletedRun>(), run => run.InnerText == "C");
        }

        [Fact]
        public void CompareStructureInPlaceRedlinesDeletedContentControlWithDelimiterText() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_deleted_control_delimiter_source.docx");
            CreateContentControlRegressionDocument(sourcePath, ("Legacy", "Legacy.Tag", "before; text=after"));

            string targetPath = Path.Combine(_directoryWithFiles, "compare_deleted_control_delimiter_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("No controlled content remains.");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_deleted_control_delimiter_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new System.Collections.Generic.HashSet<WordComparisonScope> { WordComparisonScope.ContentControl }
                    }
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.ContentControl &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            DeletedRun deleted = Assert.Single(redline._document.Body!.Descendants<DeletedRun>(), run => run.InnerText.Contains("before; text=after", StringComparison.Ordinal));
            Assert.Equal("OfficeIMO Tests", deleted.Author?.Value);
        }

        [Fact]
        public void UpdateFieldsAndGetReportAddsSeparatorForEmptyComplexFieldResults() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.EmptyComplexField.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Title = "Current title";
                document._document.Body!.Append(new Paragraph(
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" TITLE ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End })));
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                WordFieldUpdateResult update = Assert.Single(report.Results);
                Assert.Equal(WordFieldUpdateStatus.Updated, update.Status);
                Paragraph paragraph = Assert.Single(document._document.Body!.Elements<Paragraph>());
                Assert.Contains(paragraph.Elements<Run>(), run => run.Elements<FieldChar>().Any(field => field.FieldCharType?.Value == FieldCharValues.Separate));
                Assert.True(paragraph.Elements<Run>().TakeWhile(run => !run.Elements<FieldChar>().Any(field => field.FieldCharType?.Value == FieldCharValues.End)).Any(run => run.InnerText == "Current title"));
            }
        }

        [Fact]
        public void TableOfContentRefreshCaptionListAllocatesUniqueBookmarkIdsAcrossRelatedParts() {
            string filePath = Path.Combine(_directoryWithFiles, "CaptionListUniqueRelatedBookmarkIds.docx");

            using WordDocument document = WordDocument.Create(filePath);
            WordTableOfContent list = document.AddTableOfContent();
            AddGeneratedCaptionParagraph(document, "_BodyCaptionUnique", "Figure", "1", "Body deployment view");
            document.AddHeadersAndFooters();

            Header header = document._wordprocessingDocument.MainDocumentPart!.HeaderParts.Single().Header!;
            Footer footer = document._wordprocessingDocument.MainDocumentPart!.FooterParts.Single().Footer!;
            AppendCaptionParagraph(header, "Figure", "2", "Header architecture map");
            AppendCaptionParagraph(footer, "Figure", "3", "Footer recovery map");

            WordCaptionListRefreshReport report = list.RefreshListOfFigures();

            Assert.Equal(3, report.EntryCount);
            string[] bookmarkIds = document._wordprocessingDocument.MainDocumentPart!
                .HeaderParts.SelectMany(part => part.Header!.Descendants<BookmarkStart>())
                .Concat(document._wordprocessingDocument.MainDocumentPart!.FooterParts.SelectMany(part => part.Footer!.Descendants<BookmarkStart>()))
                .Concat(document._document.Body!.Descendants<BookmarkStart>())
                .Where(bookmark => bookmark.Name?.Value?.StartsWith("_OfficeIMO_Caption_", StringComparison.Ordinal) == true)
                .Select(bookmark => bookmark.Id!.Value!)
                .ToArray();

            Assert.Equal(bookmarkIds.Length, bookmarkIds.Distinct(StringComparer.Ordinal).Count());
        }

        [Fact]
        public void AcceptRevisionsStripsNestedRevisionMarkupFromPromotedRuns() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesNestedRunProperties.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document._document.Body!.Append(new Paragraph(
                    new InsertedRun(
                        new Run(
                            new RunProperties(
                                new RunPropertiesChange(
                                    new PreviousRunProperties(new Bold())) {
                                    Id = "2",
                                    Author = "OfficeIMO Tests",
                                    Date = new DateTime(2026, 6, 30, 12, 0, 0, DateTimeKind.Utc)
                                }),
                            new Text("Promoted text"))) {
                        Id = "1",
                        Author = "OfficeIMO Tests",
                        Date = new DateTime(2026, 6, 30, 12, 0, 0, DateTimeKind.Utc)
                    }));
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordRevisionOperationReport report = document.AcceptRevisions(WordRevisionFilter.All());

                Assert.Equal(2, report.MatchedRevisions.Count);
                Assert.Contains(document._document.Body!.Descendants<Run>(), run => run.InnerText == "Promoted text");
                Assert.Empty(document._document.Body!.Descendants<RunPropertiesChange>());
                Assert.Empty(document._document.Body!.Descendants<InsertedRun>());
            }
        }

        [Fact]
        public void AcceptRevisionsPreservesNonRunContentPromotedFromMoveRevision() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesMoveToHyperlink.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document._document.Body!.Append(new Paragraph(
                    new MoveToRun(
                        new Hyperlink(
                            new Run(new Text("Moved link"))) {
                            Anchor = "MovedBookmark"
                        }) {
                        Id = "1",
                        Author = "OfficeIMO Tests",
                        Date = new DateTime(2026, 6, 30, 12, 0, 0, DateTimeKind.Utc)
                    }));
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordRevisionOperationReport report = document.AcceptRevisions(WordRevisionFilter.All());

                Assert.Single(report.MatchedRevisions);
                OpenXmlElement hyperlink = Assert.Single(document._document.Body!.Descendants(), element =>
                    string.Equals(element.LocalName, "hyperlink", StringComparison.Ordinal));
                Assert.Equal("MovedBookmark", hyperlink.GetAttribute("anchor", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value);
                Assert.Equal("Moved link", hyperlink.InnerText);
                Assert.DoesNotContain(document._document.Body!.Descendants(), element =>
                    string.Equals(element.LocalName, "moveTo", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void AddCommentReservesImportedCommentParagraphIdsWithoutCommentsEx() {
            string filePath = Path.Combine(_directoryWithFiles, "CommentParaIdReservation.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Comment target");
                document.Save(false);
            }

            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true)) {
                WordprocessingCommentsPart commentsPart = package.MainDocumentPart!.AddNewPart<WordprocessingCommentsPart>();
                commentsPart.Comments = new Comments(
                    new Comment(
                        new Paragraph(
                            new Run(new Text("Imported comment")))) {
                        Id = "0",
                        Author = "Imported",
                        Initials = "IM",
                        Date = new DateTime(2026, 6, 30, 12, 0, 0, DateTimeKind.Utc)
                    });
                commentsPart.Comments.GetFirstChild<Comment>()!.GetFirstChild<Paragraph>()!.ParagraphId = "0000000A";
                commentsPart.Comments.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.AddParagraph("New comment target").AddComment("OfficeIMO Tests", "OT", "New comment");

                var commentsEx = document._wordprocessingDocument.MainDocumentPart!.WordprocessingCommentsExPart!.CommentsEx!;
                Assert.DoesNotContain(commentsEx.Elements<DocumentFormat.OpenXml.Office2013.Word.CommentEx>(), commentEx => commentEx.ParaId?.Value == "0000000A");
                Assert.Contains(commentsEx.Elements<DocumentFormat.OpenXml.Office2013.Word.CommentEx>(), commentEx => commentEx.ParaId?.Value == "0000000B");
            }
        }

        [Fact]
        public void AddCommentAllocatesParagraphIdsWithUnsignedHexOrdering() {
            string filePath = Path.Combine(_directoryWithFiles, "CommentParaIdUnsignedReservation.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Comment target");
                document.Save(false);
            }

            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true)) {
                WordprocessingCommentsPart commentsPart = package.MainDocumentPart!.AddNewPart<WordprocessingCommentsPart>();
                commentsPart.Comments = new Comments(
                    CreateImportedComment("0", "7FFFFFFF"),
                    CreateImportedComment("1", "80000000"));
                commentsPart.Comments.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.AddParagraph("New comment target").AddComment("OfficeIMO Tests", "OT", "New comment");

                var commentsEx = document._wordprocessingDocument.MainDocumentPart!.WordprocessingCommentsExPart!.CommentsEx!;
                Assert.Contains(commentsEx.Elements<DocumentFormat.OpenXml.Office2013.Word.CommentEx>(), commentEx => commentEx.ParaId?.Value == "80000001");
            }
        }

        [Fact]
        public void AddReplyUsesImportedParentParagraphIdWithoutCommentsExEntry() {
            string filePath = Path.Combine(_directoryWithFiles, "CommentReplyImportedParentParaId.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Comment target");
                document.Save(false);
            }

            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true)) {
                WordprocessingCommentsPart commentsPart = package.MainDocumentPart!.AddNewPart<WordprocessingCommentsPart>();
                commentsPart.Comments = new Comments(CreateImportedComment("0", "0000000A"));
                commentsPart.Comments.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordComment parent = Assert.Single(WordComment.GetAllComments(document));
                parent.AddReply("OfficeIMO Tests", "OT", "Imported reply");

                var commentsEx = document._wordprocessingDocument.MainDocumentPart!.WordprocessingCommentsExPart!.CommentsEx!;
                Assert.Contains(commentsEx.Elements<DocumentFormat.OpenXml.Office2013.Word.CommentEx>(), commentEx =>
                    commentEx.ParaIdParent?.Value == "0000000A");
            }
        }

        private static void CreateFieldRegressionDocument(string path, params (string Instruction, string Result)[] fields) {
            using WordDocument document = WordDocument.Create(path);
            foreach ((string instruction, string result) in fields) {
                document._document.Body!.Append(new Paragraph(
                    new Run(new Text("Field: ") { Space = SpaceProcessingModeValues.Preserve }),
                    new SimpleField(new Run(new Text(result) { Space = SpaceProcessingModeValues.Preserve })) {
                        Instruction = instruction
                    }));
            }

            document.Save(false);
        }

        private static Comment CreateImportedComment(string id, string paragraphId) {
            var comment = new Comment(
                new Paragraph(
                    new Run(new Text("Imported comment " + id)))) {
                Id = id,
                Author = "Imported",
                Initials = "IM",
                Date = new DateTime(2026, 6, 30, 12, 0, 0, DateTimeKind.Utc)
            };
            comment.GetFirstChild<Paragraph>()!.ParagraphId = paragraphId;
            return comment;
        }

        private static Paragraph CreateParagraphWithAuthorAndRawTocFields() {
            return new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" AUTHOR ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("Alice") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new Text(" before toc ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" TOC \\o \"1-3\" \\h \\z \\u ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("No table of contents entries found.") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new Text(" after toc") { Space = SpaceProcessingModeValues.Preserve }));
        }

        private static List<byte[]> GetDeletedDrawingImageBytes(OpenXmlPart part, OpenXmlElement root) {
            var images = new List<byte[]>();
            foreach (DeletedRun deletedRun in root.Descendants<DeletedRun>()) {
                foreach (DocumentFormat.OpenXml.Drawing.Blip blip in deletedRun.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()) {
                    if (blip.Embed?.Value is not string relationshipId) {
                        continue;
                    }

                    if (part.GetPartById(relationshipId) is not ImagePart imagePart) {
                        continue;
                    }

                    using Stream stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
                    using var memoryStream = new MemoryStream();
                    stream.CopyTo(memoryStream);
                    images.Add(memoryStream.ToArray());
                }
            }

            return images;
        }

        private static Table CreateSingleCellTable(string text, params OpenXmlElement[] additionalCellChildren) {
            var cell = new TableCell(
                new TableCellProperties(new TableCellWidth { Width = "5000", Type = TableWidthUnitValues.Dxa }),
                new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
            foreach (OpenXmlElement child in additionalCellChildren) {
                cell.Append(child);
            }

            return new Table(
                new TableProperties(new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto }),
                new TableGrid(new GridColumn { Width = "5000" }),
                new TableRow(cell));
        }

        private static void CreateContentControlRegressionDocument(string path, params (string Alias, string Tag, string Text)[] controls) {
            using WordDocument document = WordDocument.Create(path);
            foreach ((string alias, string tag, string text) in controls) {
                document._document.Body!.Append(new SdtBlock(
                    new SdtProperties(
                        new SdtAlias { Val = alias },
                        new Tag { Val = tag }),
                    new SdtContentBlock(
                        new Paragraph(
                            new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })))));
            }

            document.Save(false);
        }

        private static void CreateNestedContentControlDocument(string path, string outerText, string childText) {
            using WordDocument document = WordDocument.Create(path);
            document._document.Body!.Append(new SdtBlock(
                new SdtProperties(
                    new SdtAlias { Val = "Outer" },
                    new Tag { Val = "OuterTag" }),
                new SdtContentBlock(
                    new Paragraph(
                        new Run(new Text(outerText + " ") { Space = SpaceProcessingModeValues.Preserve }),
                        new SdtRun(
                            new SdtProperties(
                                new SdtAlias { Val = "Child" },
                                new Tag { Val = "ChildTag" }),
                            new SdtContentRun(
                                new Run(new Text(childText) { Space = SpaceProcessingModeValues.Preserve })))))));
            document.Save(false);
        }

        private static void CreateFootnoteIdRegressionDocument(string path, params (int Id, string Anchor, string Text)[] notes) {
            using WordDocument document = WordDocument.Create(path);
            MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;
            document._document.Body!.RemoveAllChildren<Paragraph>();

            FootnotesPart footnotesPart = mainPart.FootnotesPart ?? mainPart.AddNewPart<FootnotesPart>();
            footnotesPart.Footnotes = new Footnotes();

            foreach ((int id, string anchor, string text) in notes) {
                document._document.Body.Append(new Paragraph(
                    new Run(new Text(anchor) { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FootnoteReference { Id = id })));
                footnotesPart.Footnotes.Append(new Footnote(
                    new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }))) {
                    Id = id
                });
            }

            footnotesPart.Footnotes.Save();
            document.Save(false);
        }

        private static void AddContentControlToFirstFootnote(string path, string text) {
            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Footnote footnote = document.MainDocumentPart!.FootnotesPart!.Footnotes!.Elements<Footnote>().First(item => item.Type == null || item.Type.Value == FootnoteEndnoteValues.Normal);
            footnote.Append(new SdtBlock(
                new SdtProperties(
                    new SdtAlias { Val = "Deleted note control" },
                    new Tag { Val = "DeletedNoteControl" }),
                new SdtContentBlock(
                    new Paragraph(
                        new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })))));
            document.MainDocumentPart.FootnotesPart.Footnotes.Save();
        }

        private static void CreateRevisionRegressionDocument(string path, params string[] insertedTexts) {
            using WordDocument document = WordDocument.Create(path);
            DateTime revisionDate = new DateTime(2026, 6, 30, 12, 0, 0, DateTimeKind.Utc);
            foreach (string insertedText in insertedTexts) {
                document.AddParagraph("Tracked ").AddInsertedText(insertedText, "OfficeIMO Tests", revisionDate);
            }

            document.Save(false);
        }

        private static void ReplaceTocInstructionWithSplitComplexField(WordTableOfContent toc) {
            SimpleField field = toc.SdtBlock.Descendants<SimpleField>().First();
            Paragraph paragraph = field.Ancestors<Paragraph>().First();
            foreach (SimpleField tocField in toc.SdtBlock.Descendants<SimpleField>().ToList()) {
                tocField.Remove();
            }

            paragraph.RemoveAllChildren<Run>();
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" TO") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldCode("C \\f \"A\" \\h ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("Stale contents") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        private static void ReplaceIndexInstructionWithSplitComplexField(WordTableOfContent index) {
            SimpleField field = index.SdtBlock.Descendants<SimpleField>().First();
            Paragraph paragraph = field.Ancestors<Paragraph>().First();
            foreach (SimpleField indexField in index.SdtBlock.Descendants<SimpleField>().ToList()) {
                indexField.Remove();
            }

            paragraph.RemoveAllChildren<Run>();
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" IND") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldCode("EX \\f \"A\" ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldCode("\\c \"2\" ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("Stale index") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        private static Paragraph CreateTcFieldParagraph(string text, string entryType) {
            return new Paragraph(
                new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }),
                new SimpleField(new Run(new Text(string.Empty))) {
                    Instruction = " TC \"" + text + "\" \\f \"" + entryType + "\" "
                });
        }
    }

    public class WordFieldUpdateReportLockTests {
        private readonly string _directoryWithFiles;

        public WordFieldUpdateReportLockTests() {
            _directoryWithFiles = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TempDocuments2", Guid.NewGuid().ToString("N"));
            Word.Setup(_directoryWithFiles);
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_SkipsLockedSimpleFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.LockedSimpleField.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Title = "Updated title";
                document._document.Body!.Append(new Paragraph(
                    new Run(new Text("Title: ") { Space = SpaceProcessingModeValues.Preserve }),
                    new SimpleField(new Run(new Text("Locked title") { Space = SpaceProcessingModeValues.Preserve })) {
                        Instruction = " TITLE ",
                        FieldLock = true
                    }));
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                WordFieldUpdateResult result = Assert.Single(report.Results);
                Assert.Equal(WordFieldType.Title, result.FieldType);
                Assert.Equal(WordFieldUpdateStatus.Skipped, result.Status);
                Assert.Contains("locked", result.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Contains(document._document.Body!.Descendants<SimpleField>(), field => field.InnerText == "Locked title");
            }
        }
    }
}
