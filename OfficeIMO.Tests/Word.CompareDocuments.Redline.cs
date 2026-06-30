using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CompareStructureCreatesRedlineDocumentWithTrackedChangesAndReport() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Status: Draft");
                document.AddParagraph("Remove this clause.");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Status: Approved");
                document.AddParagraph("Add this clause.");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 28, 12, 0, 0, DateTimeKind.Utc)
                });

            Assert.True(File.Exists(outputPath));
            Assert.True(result.HasChanges);
            Assert.Contains(result.Findings, finding => finding.SourceText == "Status: Draft" && finding.TargetText == "Status: Approved");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Assert.Contains(body.Descendants<InsertedRun>(), run => run.InnerText == "Status: Approved" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(body.Descendants<DeletedRun>(), run => run.InnerText == "Status: Draft" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline"));
            Assert.Contains(redline.Paragraphs, paragraph => paragraph.Text.Contains("Tracked Changes"));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureRedlineCanKeepFeatureFindingsReportOnly() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_feature_policy_source.docx");
            CreateDocumentWithSimpleField(sourcePath, " AUTHOR ", "Same result");
            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_feature_policy_target.docx");
            CreateDocumentWithSimpleField(targetPath, " TITLE ", "Same result");

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_feature_policy_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    TrackFeatureFindings = false
                });

            Assert.Contains(result.Findings, finding => finding.Scope == WordComparisonScope.Field);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Assert.DoesNotContain(body.Descendants<InsertedRun>(), run => run.InnerText.Contains("TITLE", StringComparison.Ordinal));
            Assert.DoesNotContain(body.Descendants<DeletedRun>(), run => run.InnerText.Contains("AUTHOR", StringComparison.Ordinal));
            Assert.Contains(redline.Paragraphs, paragraph => paragraph.Text.Contains("Field", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureRedlineTracksFeatureFindingsWhenTextFindingsAreDisabled() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_feature_without_text_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                AddNestedRunContentControl(document, "Contoso");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_feature_without_text_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                AddNestedRunContentControl(document, "Fabrikam");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_feature_without_text_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    TrackTextFindings = false,
                    TrackFeatureFindings = true
                });

            Assert.Contains(result.Findings, finding => finding.Scope == WordComparisonScope.ContentControl);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            SdtRun innerControl = body.Descendants<SdtRun>().Last();
            Assert.Contains(innerControl.Descendants<InsertedRun>(), run => run.InnerText == "Fabrikam" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(innerControl.Descendants<DeletedRun>(), run => run.InnerText == "Contoso" && run.Author?.Value == "OfficeIMO Tests");
        }

        [Fact]
        public void CompareStructureRedlineCanKeepReviewAndFormattingFindingsReportOnly() {
            string reviewSourcePath = Path.Combine(_directoryWithFiles, "compare_redline_review_policy_source.docx");
            using (WordDocument document = WordDocument.Create(reviewSourcePath)) {
                document.AddParagraph("Review target").AddComment("Alice Reviewer", "AR", "Source note.");
                document.Save(false);
            }

            string reviewTargetPath = Path.Combine(_directoryWithFiles, "compare_redline_review_policy_target.docx");
            using (WordDocument document = WordDocument.Create(reviewTargetPath)) {
                document.AddParagraph("Review target").AddComment("Alice Reviewer", "AR", "Target note.");
                document.Save(false);
            }

            string reviewOutputPath = Path.Combine(_directoryWithFiles, "compare_redline_review_policy_output.docx");
            WordComparisonResult reviewResult = WordDocumentComparer.CreateRedlineDocument(
                reviewSourcePath,
                reviewTargetPath,
                reviewOutputPath,
                new WordComparisonRedlineOptions {
                    TrackReviewFindings = false
                });
            Assert.Contains(reviewResult.Findings, finding => finding.Scope == WordComparisonScope.Comment);

            using (WordDocument redline = WordDocument.Load(reviewOutputPath, readOnly: true)) {
                Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
                Assert.Empty(body.Descendants<InsertedRun>());
                Assert.Empty(body.Descendants<DeletedRun>());
                Assert.Contains(redline.Paragraphs, paragraph => paragraph.Text.Contains("Comment", StringComparison.Ordinal));
            }

            string formatSourcePath = Path.Combine(_directoryWithFiles, "compare_redline_format_policy_source.docx");
            CreateDocumentWithSingleRun(formatSourcePath, bold: false);
            string formatTargetPath = Path.Combine(_directoryWithFiles, "compare_redline_format_policy_target.docx");
            CreateDocumentWithSingleRun(formatTargetPath, bold: true);

            string formatOutputPath = Path.Combine(_directoryWithFiles, "compare_redline_format_policy_output.docx");
            WordComparisonResult formatResult = WordDocumentComparer.CreateRedlineDocument(
                formatSourcePath,
                formatTargetPath,
                formatOutputPath,
                new WordComparisonRedlineOptions {
                    TrackFormattingFindings = false
                });
            Assert.Contains(formatResult.Findings, finding =>
                finding.Scope == WordComparisonScope.Run &&
                finding.Message == "Run formatting changed.");

            using (WordDocument redline = WordDocument.Load(formatOutputPath, readOnly: true)) {
                Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
                Assert.Empty(body.Descendants<InsertedRun>());
                Assert.Empty(body.Descendants<DeletedRun>());
                Assert.Contains(redline.Paragraphs, paragraph => paragraph.Text.Contains("Run formatting changed.", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void CompareStructureRedlineTracksReviewFindingsWhenFeatureFindingsAreDisabled() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_review_without_feature_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Review target").AddComment("Alice Reviewer", "AR", "Source note.");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_review_without_feature_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Review target").AddComment("Alice Reviewer", "AR", "Target note.");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_review_without_feature_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    TrackFeatureFindings = false,
                    TrackReviewFindings = true
                });

            Assert.Contains(result.Findings, finding => finding.Scope == WordComparisonScope.Comment);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Assert.Contains(body.Descendants<InsertedRun>(), run => run.InnerText.Contains("Target note.", StringComparison.Ordinal));
            Assert.Contains(body.Descendants<DeletedRun>(), run => run.InnerText.Contains("Source note.", StringComparison.Ordinal));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForRunFormattingChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_run_format_source.docx");
            CreateDocumentWithSingleRun(sourcePath, bold: false);
            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_run_format_target.docx");
            CreateDocumentWithSingleRun(targetPath, bold: true);

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_run_format_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 30, 10, 0, 0, DateTimeKind.Utc)
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Run &&
                finding.Message == "Run formatting changed.");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Assert.Empty(body.Descendants<InsertedRun>());
            Assert.Empty(body.Descendants<DeletedRun>());
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            Run run = body.Descendants<Run>().Single();
            Assert.NotNull(run.RunProperties?.Bold);
            RunPropertiesChange change = Assert.Single(run.RunProperties!.Elements<RunPropertiesChange>());
            Assert.Equal("OfficeIMO Tests", change.Author?.Value);
            Assert.NotNull(change.PreviousRunProperties);
            Assert.Empty(change.PreviousRunProperties!.Elements<Bold>());

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForBodyParagraphChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Quarterly report").SetStyle(WordParagraphStyles.Heading1);
                document.AddParagraph("Status: Draft");
                document.AddParagraph("Closing note.");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Quarterly report").SetStyle(WordParagraphStyles.Heading1);
                document.AddParagraph("Status: Approved");
                document.AddParagraph("Closing note.");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 12, 0, 0, DateTimeKind.Utc)
                });

            Assert.True(result.HasChanges);
            Assert.Contains(result.Findings, finding => finding.SourceText == "Status: Draft" && finding.TargetText == "Status: Approved");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));
            Assert.Contains(redline.Paragraphs, paragraph => paragraph.Text.Contains("Quarterly report", StringComparison.Ordinal));
            Assert.Contains(body.Descendants<DeletedRun>(), run => run.InnerText == "Status: Draft" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(body.Descendants<InsertedRun>(), run => run.InnerText == "Status: Approved" && run.Author?.Value == "OfficeIMO Tests");

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForDeletedBodyParagraphs() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_deleted_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Keep before.");
                document.AddParagraph("Delete this clause.");
                document.AddParagraph("Keep after.");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_deleted_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Keep before.");
                document.AddParagraph("Keep after.");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_deleted_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding => finding.ChangeKind == WordComparisonChangeKind.Deleted && finding.SourceText == "Delete this clause.");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Assert.Contains(body.Descendants<DeletedRun>(), run => run.InnerText == "Delete this clause." && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(redline.Paragraphs, paragraph => paragraph.Text.Contains("Keep before.", StringComparison.Ordinal));
            Assert.Contains(redline.Paragraphs, paragraph => paragraph.Text.Contains("Keep after.", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForHeaderAndFooterParagraphChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_header_footer_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Body stays stable.");
                document.HeaderDefaultOrCreate.AddParagraph("Classification: Draft");
                document.FooterDefaultOrCreate.AddParagraph("Footer note: Internal");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_header_footer_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Body stays stable.");
                document.HeaderDefaultOrCreate.AddParagraph("Classification: Final");
                document.FooterDefaultOrCreate.AddParagraph("Footer note: Published");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_header_footer_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 14, 0, 0, DateTimeKind.Utc)
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.SourceText == "Classification: Draft" &&
                finding.TargetText == "Classification: Final");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.SourceText == "Footer note: Internal" &&
                finding.TargetText == "Footer note: Published");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Header header = redline._wordprocessingDocument.MainDocumentPart!.HeaderParts.Single().Header!;
            Footer footer = redline._wordprocessingDocument.MainDocumentPart!.FooterParts.Single().Footer!;

            Assert.Contains(header.Descendants<DeletedRun>(), run => run.InnerText == "Classification: Draft" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(header.Descendants<InsertedRun>(), run => run.InnerText == "Classification: Final" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(footer.Descendants<DeletedRun>(), run => run.InnerText == "Footer note: Internal" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(footer.Descendants<InsertedRun>(), run => run.InnerText == "Footer note: Published" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!.Descendants<InsertedRun>(), run => run.InnerText.Contains("Classification", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForHeaderTableCellChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_header_table_cell_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Body stays stable.");
                WordTable table = document.HeaderDefaultOrCreate.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Classification";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Draft";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Legal";
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_header_table_cell_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Body stays stable.");
                WordTable table = document.HeaderDefaultOrCreate.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Classification";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Final";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Legal";
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_header_table_cell_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.SourceText == "Draft" &&
                finding.TargetText == "Final");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Header header = redline._wordprocessingDocument.MainDocumentPart!.HeaderParts.Single().Header!;
            TableCell changedCell = header.Descendants<Table>().Single().Elements<TableRow>().ElementAt(1).Elements<TableCell>().ElementAt(0);
            Assert.Contains(changedCell.Descendants<DeletedRun>(), run => run.InnerText == "Draft" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(changedCell.Descendants<InsertedRun>(), run => run.InnerText == "Final" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!.Descendants<InsertedRun>(), run => run.InnerText == "Final");

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForFooterInsertedTableRows() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_footer_table_row_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Body stays stable.");
                WordTable table = document.FooterDefaultOrCreate.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Retention";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Legal";
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_footer_table_row_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Body stays stable.");
                WordTable table = document.FooterDefaultOrCreate.AddTable(3, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Escalation";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Support";
                table.Rows[2].Cells[0].Paragraphs[0].Text = "Retention";
                table.Rows[2].Cells[1].Paragraphs[0].Text = "Legal";
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_footer_table_row_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableRow &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Escalation | Support");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Footer footer = redline._wordprocessingDocument.MainDocumentPart!.FooterParts.Single().Footer!;
            TableRow insertedRow = footer.Descendants<Table>().Single().Elements<TableRow>().ElementAt(1);
            Assert.Contains(insertedRow.Elements<TableCell>().ElementAt(0).Descendants<InsertedRun>(), run => run.InnerText == "Escalation" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(insertedRow.Elements<TableCell>().ElementAt(1).Descendants<InsertedRun>(), run => run.InnerText == "Support" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!.Descendants<InsertedRun>(), run => run.InnerText == "Escalation");

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForDeletedFooterTables() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_footer_table_deleted_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Body stays stable.");
                WordTable table = document.FooterDefaultOrCreate.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Legacy";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Operations";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Archive";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Annual";
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_footer_table_deleted_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Body stays stable.");
                document.FooterDefaultOrCreate.AddParagraph("Footer remains.");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_footer_table_deleted_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Table &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted &&
                finding.SourceText != null &&
                finding.SourceText.Contains("Legacy", StringComparison.Ordinal));

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Footer footer = redline._wordprocessingDocument.MainDocumentPart!.FooterParts.Single().Footer!;
            Table deletedTable = Assert.Single(footer.Descendants<Table>());
            Assert.Contains(deletedTable.Descendants<DeletedRun>(), run => run.InnerText == "Legacy" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(deletedTable.Descendants<DeletedRun>(), run => run.InnerText == "Annual" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Empty(body.Descendants<Table>());

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForFootnoteAndEndnoteParagraphChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_notes_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Body footnote anchor").AddFootNote("Source footnote text");
                document.AddParagraph("Body endnote anchor").AddEndNote("Source endnote text");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_notes_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Body footnote anchor").AddFootNote("Target footnote text");
                document.AddParagraph("Body endnote anchor").AddEndNote("Target endnote text");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_notes_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 15, 0, 0, DateTimeKind.Utc)
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.SourceText == "Source footnote text" &&
                finding.TargetText == "Target footnote text");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.SourceText == "Source endnote text" &&
                finding.TargetText == "Target endnote text");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Footnotes footnotes = redline._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!;
            Endnotes endnotes = redline._wordprocessingDocument.MainDocumentPart!.EndnotesPart!.Endnotes!;

            Assert.Contains(footnotes.Descendants<DeletedRun>(), run => run.InnerText == "Source footnote text" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(footnotes.Descendants<InsertedRun>(), run => run.InnerText == "Target footnote text" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(endnotes.Descendants<DeletedRun>(), run => run.InnerText == "Source endnote text" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(endnotes.Descendants<InsertedRun>(), run => run.InnerText == "Target endnote text" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(body.Descendants<DeletedRun>(), run => run.InnerText.Contains("footnote", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(body.Descendants<InsertedRun>(), run => run.InnerText.Contains("endnote", StringComparison.OrdinalIgnoreCase));
            Assert.NotEmpty(body.Descendants<FootnoteReference>());
            Assert.NotEmpty(body.Descendants<EndnoteReference>());

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForFootnoteTableCellChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_footnote_table_cell_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Body footnote table anchor").AddFootNote("Source footnote table");
                ReplaceLastFootnoteWithTable(document, new[] { "Control", "Owner" }, new[] { "Retention", "Legal" });
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_footnote_table_cell_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Body footnote table anchor").AddFootNote("Target footnote table");
                ReplaceLastFootnoteWithTable(document, new[] { "Control", "Owner" }, new[] { "Retention", "Compliance" });
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_footnote_table_cell_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.SourceText == "Legal" &&
                finding.TargetText == "Compliance");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Footnotes footnotes = redline._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!;
            TableCell changedCell = footnotes.Descendants<Table>().Single().Elements<TableRow>().ElementAt(1).Elements<TableCell>().ElementAt(1);
            Assert.Contains(changedCell.Descendants<DeletedRun>(), run => run.InnerText == "Legal" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(changedCell.Descendants<InsertedRun>(), run => run.InnerText == "Compliance" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!.Descendants<InsertedRun>(), run => run.InnerText == "Compliance");

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForEndnoteInsertedTableRows() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_endnote_table_row_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Body endnote table anchor").AddEndNote("Source endnote table");
                ReplaceLastEndnoteWithTable(document, new[] { "Control", "Owner" }, new[] { "Retention", "Legal" });
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_endnote_table_row_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Body endnote table anchor").AddEndNote("Target endnote table");
                ReplaceLastEndnoteWithTable(document, new[] { "Control", "Owner" }, new[] { "Escalation", "Support" }, new[] { "Retention", "Legal" });
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_endnote_table_row_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableRow &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Escalation | Support");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Endnotes endnotes = redline._wordprocessingDocument.MainDocumentPart!.EndnotesPart!.Endnotes!;
            TableRow insertedRow = endnotes.Descendants<Table>().Single().Elements<TableRow>().ElementAt(1);
            Assert.Contains(insertedRow.Elements<TableCell>().ElementAt(0).Descendants<InsertedRun>(), run => run.InnerText == "Escalation" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(insertedRow.Elements<TableCell>().ElementAt(1).Descendants<InsertedRun>(), run => run.InnerText == "Support" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!.Descendants<InsertedRun>(), run => run.InnerText == "Escalation");

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForDeletedFootnoteTables() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_footnote_table_deleted_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Body deleted footnote table anchor").AddFootNote("Source footnote table");
                ReplaceLastFootnoteWithTable(document, new[] { "Legacy", "Operations" }, new[] { "Archive", "Annual" });
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_footnote_table_deleted_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Body deleted footnote table anchor").AddFootNote("Target footnote without table");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_footnote_table_deleted_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Table &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted &&
                finding.SourceText != null &&
                finding.SourceText.Contains("Legacy", StringComparison.Ordinal));

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Footnotes footnotes = redline._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!;
            Table deletedTable = Assert.Single(footnotes.Descendants<Table>());
            Assert.Contains(deletedTable.Descendants<DeletedRun>(), run => run.InnerText == "Legacy" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(deletedTable.Descendants<DeletedRun>(), run => run.InnerText == "Annual" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Empty(body.Descendants<Table>());

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForContentControlTextChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_content_control_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                WordParagraph paragraph = document.AddParagraph("Client: ");
                paragraph.AddStructuredDocumentTag("Contoso", "Client", "ClientName");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_content_control_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                WordParagraph paragraph = document.AddParagraph("Client: ");
                paragraph.AddStructuredDocumentTag("Fabrikam", "Client", "ClientName");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_content_control_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 16, 0, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.ContentControl
                        }
                    }
                });

            WordComparisonFinding contentControlFinding = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.ContentControl &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Contains("text=Contoso", contentControlFinding.SourceText, StringComparison.Ordinal);
            Assert.Contains("text=Fabrikam", contentControlFinding.TargetText, StringComparison.Ordinal);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            SdtRun contentControl = Assert.Single(body.Descendants<SdtRun>());
            Assert.Contains(contentControl.Descendants<DeletedRun>(), run => run.InnerText == "Contoso" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(contentControl.Descendants<InsertedRun>(), run => run.InnerText == "Fabrikam" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(body.Descendants<Paragraph>(), paragraph => paragraph.InnerText.Contains("Client: ", StringComparison.Ordinal));
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForTextBoxContentControlTextChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_textbox_content_control_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                AddTextBoxRunContentControl(document, "Pending");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_textbox_content_control_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                AddTextBoxRunContentControl(document, "Approved");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_textbox_content_control_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 18, 10, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.ContentControl
                        }
                    }
                });

            WordComparisonFinding contentControlFinding = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.ContentControl &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Contains("text-box", contentControlFinding.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("text=Pending", contentControlFinding.SourceText, StringComparison.Ordinal);
            Assert.Contains("text=Approved", contentControlFinding.TargetText, StringComparison.Ordinal);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            TextBoxContent textBoxContent = Assert.Single(body.Descendants<TextBoxContent>());
            SdtRun contentControl = Assert.Single(textBoxContent.Descendants<SdtRun>());
            Assert.Contains(contentControl.Descendants<DeletedRun>(), run => run.InnerText == "Pending" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(contentControl.Descendants<InsertedRun>(), run => run.InnerText == "Approved" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForTextBoxBlockContentControlTextChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_textbox_block_content_control_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                AddTextBoxBlockContentControl(document, "Legal review pending");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_textbox_block_content_control_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                AddTextBoxBlockContentControl(document, "Legal review approved");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_textbox_block_content_control_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 18, 20, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.ContentControl
                        }
                    }
                });

            WordComparisonFinding contentControlFinding = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.ContentControl &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Contains("text-box", contentControlFinding.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("content-control[0]", contentControlFinding.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("text=Legal review pending", contentControlFinding.SourceText, StringComparison.Ordinal);
            Assert.Contains("text=Legal review approved", contentControlFinding.TargetText, StringComparison.Ordinal);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            TextBoxContent textBoxContent = Assert.Single(body.Descendants<TextBoxContent>());
            SdtBlock contentControl = Assert.Single(textBoxContent.Descendants<SdtBlock>());
            Assert.Contains(contentControl.Descendants<DeletedRun>(), run => run.InnerText == "Legal review pending" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(contentControl.Descendants<InsertedRun>(), run => run.InnerText == "Legal review approved" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForTableContentControlTextChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_content_control_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                WordTable sourceTable = document.AddTable(1, 1);
                sourceTable.Rows[0].Cells[0].Paragraphs[0].AddText("Client: ");
                sourceTable.Rows[0].Cells[0].Paragraphs[0].AddStructuredDocumentTag("Contoso", "Client", "ClientName");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_content_control_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                WordTable targetTable = document.AddTable(1, 1);
                targetTable.Rows[0].Cells[0].Paragraphs[0].AddText("Client: ");
                targetTable.Rows[0].Cells[0].Paragraphs[0].AddStructuredDocumentTag("Fabrikam", "Client", "ClientName");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_content_control_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 16, 30, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.ContentControl
                        }
                    }
                });

            WordComparisonFinding contentControlFinding = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.ContentControl &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Contains("table", contentControlFinding.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("text=Contoso", contentControlFinding.SourceText, StringComparison.Ordinal);
            Assert.Contains("text=Fabrikam", contentControlFinding.TargetText, StringComparison.Ordinal);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Table rawTable = Assert.Single(body.Descendants<Table>());
            TableCell cell = Assert.Single(rawTable.Descendants<TableCell>());
            SdtRun contentControl = Assert.Single(cell.Descendants<SdtRun>());
            Assert.Contains(contentControl.Descendants<DeletedRun>(), run => run.InnerText == "Contoso" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(contentControl.Descendants<InsertedRun>(), run => run.InnerText == "Fabrikam" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(cell.Descendants<Paragraph>(), paragraph => paragraph.InnerText.Contains("Client: ", StringComparison.Ordinal));
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForTableBlockContentControlTextChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_block_content_control_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                WordTable sourceTable = document.AddTable(1, 1);
                ReplaceCellWithBlockContentControl(sourceTable.Rows[0].Cells[0], "Evidence pending");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_block_content_control_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                WordTable targetTable = document.AddTable(1, 1);
                ReplaceCellWithBlockContentControl(targetTable.Rows[0].Cells[0], "Evidence approved");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_block_content_control_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 16, 45, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.ContentControl
                        }
                    }
                });

            WordComparisonFinding contentControlFinding = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.ContentControl &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Contains("content-control[0]", contentControlFinding.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("table", contentControlFinding.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("text=Evidence pending", contentControlFinding.SourceText, StringComparison.Ordinal);
            Assert.Contains("text=Evidence approved", contentControlFinding.TargetText, StringComparison.Ordinal);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Table rawTable = Assert.Single(body.Descendants<Table>());
            TableCell cell = Assert.Single(rawTable.Descendants<TableCell>());
            SdtBlock contentControl = Assert.Single(cell.Descendants<SdtBlock>());
            Assert.Contains(contentControl.Descendants<DeletedRun>(), run => run.InnerText == "Evidence pending" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(contentControl.Descendants<InsertedRun>(), run => run.InnerText == "Evidence approved" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForTableCellContentControlTextChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_cell_sdt_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                WordTable sourceTable = document.AddTable(1, 1);
                sourceTable.Rows[0].Cells[0].Paragraphs[0].Text = "Cell pending";
                WrapCellInCellContentControl(sourceTable.Rows[0].Cells[0], "CellStatus");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_cell_sdt_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                WordTable targetTable = document.AddTable(1, 1);
                targetTable.Rows[0].Cells[0].Paragraphs[0].Text = "Cell approved";
                WrapCellInCellContentControl(targetTable.Rows[0].Cells[0], "CellStatus");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_cell_sdt_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 17, 10, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.ContentControl
                        }
                    }
                });

            WordComparisonFinding contentControlFinding = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.ContentControl &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Contains("content-control[0]", contentControlFinding.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("table", contentControlFinding.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("text=Cell pending", contentControlFinding.SourceText, StringComparison.Ordinal);
            Assert.Contains("text=Cell approved", contentControlFinding.TargetText, StringComparison.Ordinal);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Table rawTable = Assert.Single(body.Descendants<Table>());
            SdtCell contentControl = Assert.Single(rawTable.Descendants<SdtCell>());
            TableCell cell = Assert.Single(contentControl.Descendants<TableCell>());
            Assert.Contains(cell.Descendants<DeletedRun>(), run => run.InnerText == "Cell pending" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(cell.Descendants<InsertedRun>(), run => run.InnerText == "Cell approved" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForTableRowContentControlTextChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_row_sdt_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                WordTable sourceTable = document.AddTable(1, 1);
                sourceTable.Rows[0].Cells[0].Paragraphs[0].Text = "Row pending";
                WrapRowInRowContentControl(sourceTable.Rows[0], "RowStatus");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_row_sdt_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                WordTable targetTable = document.AddTable(1, 1);
                targetTable.Rows[0].Cells[0].Paragraphs[0].Text = "Row approved";
                WrapRowInRowContentControl(targetTable.Rows[0], "RowStatus");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_row_sdt_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 17, 20, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.ContentControl
                        }
                    }
                });

            WordComparisonFinding contentControlFinding = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.ContentControl &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Contains("content-control[0]", contentControlFinding.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("table", contentControlFinding.DetailedLocation, StringComparison.Ordinal);
            Assert.Contains("text=Row pending", contentControlFinding.SourceText, StringComparison.Ordinal);
            Assert.Contains("text=Row approved", contentControlFinding.TargetText, StringComparison.Ordinal);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Table rawTable = Assert.Single(body.Descendants<Table>());
            SdtRow contentControl = Assert.Single(rawTable.Descendants<SdtRow>());
            TableRow row = Assert.Single(contentControl.Descendants<TableRow>());
            Assert.Contains(row.Descendants<DeletedRun>(), run => run.InnerText == "Row pending" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(row.Descendants<InsertedRun>(), run => run.InnerText == "Row approved" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForNestedContentControlTextChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_nested_content_control_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                AddNestedRunContentControl(document, "Contoso");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_nested_content_control_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                AddNestedRunContentControl(document, "Fabrikam");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_nested_content_control_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 17, 30, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.ContentControl
                        }
                    }
                });

            List<WordComparisonFinding> contentControlFindings = result.Findings
                .Where(finding => finding.Scope == WordComparisonScope.ContentControl &&
                                  finding.ChangeKind == WordComparisonChangeKind.Modified)
                .ToList();
            Assert.True(contentControlFindings.Count >= 2);
            Assert.Contains(contentControlFindings, finding =>
                finding.DetailedLocation.Contains("nested-content-control", StringComparison.Ordinal) &&
                finding.SourceText.Contains("text=Contoso", StringComparison.Ordinal) &&
                finding.TargetText.Contains("text=Fabrikam", StringComparison.Ordinal));

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            SdtRun outerControl = Assert.Single(body.Descendants<SdtRun>(), control => !control.Ancestors<SdtRun>().Any());
            SdtRun innerControl = Assert.Single(outerControl.Descendants<SdtRun>());
            Assert.Contains(outerControl.SdtContentRun!.Elements<Run>(), run => run.InnerText == "Client: ");
            Assert.Contains(innerControl.Descendants<DeletedRun>(), run => run.InnerText == "Contoso" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(innerControl.Descendants<InsertedRun>(), run => run.InnerText == "Fabrikam" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(outerControl.SdtContentRun!.Elements<DeletedRun>(), run => run.InnerText == "Contoso");
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureInPlaceParagraphRedlinePreservesNonTextRuns() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_mixed_paragraph_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document._document.Body!.Append(new Paragraph(
                    new Run(new Text("Status: Draft") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(CreateNonImageDrawing())));
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_mixed_paragraph_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document._document.Body!.Append(new Paragraph(
                    new Run(new Text("Status: Approved") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(CreateNonImageDrawing())));
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_mixed_paragraph_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Paragraph paragraph = Assert.Single(body.Elements<Paragraph>());
            Assert.Contains(paragraph.Descendants<DeletedRun>(), run => run.InnerText == "Status: Draft");
            Assert.Contains(paragraph.Descendants<InsertedRun>(), run => run.InnerText == "Status: Approved");
            Assert.Single(paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForInsertedImages() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_image_inserted_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Before image");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_image_inserted_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Before image");
                document.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"), 80, 40);
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_image_inserted_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 17, 0, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.Image
                        }
                    }
                });

            WordComparisonFinding imageFinding = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted);
            Assert.Equal("image[0]", imageFinding.Location);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            InsertedRun insertedImage = Assert.Single(body.Descendants<InsertedRun>(), run => run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any());
            Assert.Equal("OfficeIMO Tests", insertedImage.Author?.Value);
            Assert.NotEmpty(insertedImage.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
            Assert.DoesNotContain(body.Elements<Paragraph>().SelectMany(paragraph => paragraph.Elements<Run>()), run => run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any());
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureImageRedlineSkipsNonImageDrawingsWhenMappingIndexes() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_image_non_image_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Before drawing");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_image_non_image_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Before drawing");
                document._document.Body!.Append(new Paragraph(new Run(CreateNonImageDrawing())));
                document.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"), 80, 40);
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_image_non_image_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.Image
                        }
                    }
                });

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            InsertedRun insertedImage = Assert.Single(body.Descendants<InsertedRun>(), run => run.Descendants<A.Blip>().Any());
            Assert.Equal("OfficeIMO Tests", insertedImage.Author?.Value);
            Assert.DoesNotContain(body.Descendants<InsertedRun>(), run =>
                run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any() &&
                !run.Descendants<A.Blip>().Any());
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForDeletedImages() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_image_deleted_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Before image");
                document.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"), 80, 40);
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_image_deleted_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Before image");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_image_deleted_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 17, 15, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.Image
                        }
                    }
                });

            WordComparisonFinding imageFinding = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted);
            Assert.Equal("image[0]", imageFinding.Location);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            DeletedRun deletedImage = Assert.Single(body.Descendants<DeletedRun>(), run => run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any());
            Assert.Equal("OfficeIMO Tests", deletedImage.Author?.Value);
            Assert.NotEmpty(deletedImage.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
            Assert.DoesNotContain(body.Descendants<InsertedRun>(), run => run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any());
            Assert.NotEmpty(redline._wordprocessingDocument.MainDocumentPart!.ImageParts);
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForChangedImagePayloads() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_image_changed_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Before image");
                document.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"), 80, 40);
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_image_changed_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Before image");
                document.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"), 80, 40);
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_image_changed_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 17, 30, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.Image
                        }
                    }
                });

            WordComparisonFinding imageFinding = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("Image payload changed.", imageFinding.Message);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            DeletedRun deletedImage = Assert.Single(body.Descendants<DeletedRun>(), run => run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any());
            InsertedRun insertedImage = Assert.Single(body.Descendants<InsertedRun>(), run => run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any());
            Assert.Equal("OfficeIMO Tests", deletedImage.Author?.Value);
            Assert.Equal("OfficeIMO Tests", insertedImage.Author?.Value);
            Assert.NotEqual(
                deletedImage.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Single().Embed?.Value,
                insertedImage.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Single().Embed?.Value);
            Assert.DoesNotContain(body.Elements<Paragraph>().SelectMany(paragraph => paragraph.Elements<Run>()), run => run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any());
            Assert.True(redline._wordprocessingDocument.MainDocumentPart!.ImageParts.Count() >= 2);
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForInsertedVmlImages() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_vml_image_inserted_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Before image");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_vml_image_inserted_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Before image");
                document.AddParagraph().AddImageVml(Path.Combine(_directoryWithImages, "EvotecLogo.png"), 80, 40);
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_vml_image_inserted_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 17, 45, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.Image
                        }
                    }
                });

            WordComparisonFinding imageFinding = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted);
            Assert.Equal("image[0]", imageFinding.Location);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            InsertedRun insertedImage = Assert.Single(body.Descendants<InsertedRun>(), run => run.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().Any());
            Assert.Equal("OfficeIMO Tests", insertedImage.Author?.Value);
            Assert.NotEmpty(insertedImage.Descendants<DocumentFormat.OpenXml.Vml.ImageData>());
            Assert.DoesNotContain(body.Elements<Paragraph>().SelectMany(paragraph => paragraph.Elements<Run>()), run => run.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().Any());
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForDeletedVmlImages() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_vml_image_deleted_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Before image");
                document.AddParagraph().AddImageVml(Path.Combine(_directoryWithImages, "EvotecLogo.png"), 80, 40);
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_vml_image_deleted_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Before image");
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_vml_image_deleted_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 18, 0, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.Image
                        }
                    }
                });

            WordComparisonFinding imageFinding = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted);
            Assert.Equal("image[0]", imageFinding.Location);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            DeletedRun deletedImage = Assert.Single(body.Descendants<DeletedRun>(), run => run.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().Any());
            Assert.Equal("OfficeIMO Tests", deletedImage.Author?.Value);
            Assert.NotEmpty(deletedImage.Descendants<DocumentFormat.OpenXml.Vml.ImageData>());
            Assert.DoesNotContain(body.Descendants<InsertedRun>(), run => run.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().Any());
            Assert.NotEmpty(redline._wordprocessingDocument.MainDocumentPart!.ImageParts);
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForChangedVmlImagePayloads() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_vml_image_changed_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Before image");
                document.AddParagraph().AddImageVml(Path.Combine(_directoryWithImages, "EvotecLogo.png"), 80, 40);
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_vml_image_changed_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Before image");
                document.AddParagraph().AddImageVml(Path.Combine(_directoryWithImages, "Kulek.jpg"), 80, 40);
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_vml_image_changed_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 18, 15, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.Image
                        }
                    }
                });

            WordComparisonFinding imageFinding = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Image &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("Image payload changed.", imageFinding.Message);

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            DeletedRun deletedImage = Assert.Single(body.Descendants<DeletedRun>(), run => run.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().Any());
            InsertedRun insertedImage = Assert.Single(body.Descendants<InsertedRun>(), run => run.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().Any());
            Assert.Equal("OfficeIMO Tests", deletedImage.Author?.Value);
            Assert.Equal("OfficeIMO Tests", insertedImage.Author?.Value);
            Assert.NotEqual(
                deletedImage.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().Single().RelationshipId?.Value,
                insertedImage.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().Single().RelationshipId?.Value);
            Assert.DoesNotContain(body.Elements<Paragraph>().SelectMany(paragraph => paragraph.Elements<Run>()), run => run.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().Any());
            Assert.True(redline._wordprocessingDocument.MainDocumentPart!.ImageParts.Count() >= 2);
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForModifiedTableCells() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                WordTable table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Retention";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Legal";
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                WordTable table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Retention";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Compliance";
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 6, 29, 13, 0, 0, DateTimeKind.Utc)
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.SourceText == "Legal" &&
                finding.TargetText == "Compliance");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            TableCell changedCell = body.Descendants<Table>().First().Elements<TableRow>().ElementAt(1).Elements<TableCell>().ElementAt(1);
            Assert.Contains(changedCell.Descendants<DeletedRun>(), run => run.InnerText == "Legal" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(changedCell.Descendants<InsertedRun>(), run => run.InnerText == "Compliance" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(redline.Paragraphs, paragraph => paragraph.Text.Contains("Word Comparison Redline", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForDeletedTableCells() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_deleted_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                WordTable table = document.AddTable(1, 3);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Deprecated";
                table.Rows[0].Cells[2].Paragraphs[0].Text = "Owner";
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_deleted_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                WordTable table = document.AddTable(1, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_deleted_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted &&
                finding.SourceText == "Deprecated");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            TableRow row = body.Descendants<Table>().First().Elements<TableRow>().First();
            Assert.Equal(3, row.Elements<TableCell>().Count());
            Assert.Contains(row.Elements<TableCell>().ElementAt(1).Descendants<DeletedRun>(), run => run.InnerText == "Deprecated" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(row.Elements<TableCell>().ElementAt(2).InnerText, "Owner", StringComparison.Ordinal);

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForInsertedTableCells() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_inserted_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                WordTable table = document.AddTable(1, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_inserted_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                WordTable table = document.AddTable(1, 3);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Priority";
                table.Rows[0].Cells[2].Paragraphs[0].Text = "Owner";
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_inserted_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Priority");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            TableRow row = body.Descendants<Table>().First().Elements<TableRow>().First();
            Assert.Equal(3, row.Elements<TableCell>().Count());
            Assert.Contains(row.Elements<TableCell>().ElementAt(1).Descendants<InsertedRun>(), run => run.InnerText == "Priority" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(row.Elements<TableCell>().ElementAt(2).InnerText, "Owner", StringComparison.Ordinal);

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForInsertedTableRows() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_row_inserted_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                WordTable table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Retention";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Legal";
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_row_inserted_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                WordTable table = document.AddTable(3, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Escalation";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Support";
                table.Rows[2].Cells[0].Paragraphs[0].Text = "Retention";
                table.Rows[2].Cells[1].Paragraphs[0].Text = "Legal";
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_row_inserted_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableRow &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText == "Escalation | Support");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            TableRow insertedRow = body.Descendants<Table>().First().Elements<TableRow>().ElementAt(1);
            Assert.Contains(insertedRow.Elements<TableCell>().ElementAt(0).Descendants<InsertedRun>(), run => run.InnerText == "Escalation" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(insertedRow.Elements<TableCell>().ElementAt(1).Descendants<InsertedRun>(), run => run.InnerText == "Support" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains("Retention", body.Descendants<Table>().First().Elements<TableRow>().ElementAt(2).InnerText, StringComparison.Ordinal);

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForDeletedTableRows() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_row_deleted_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                WordTable table = document.AddTable(3, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Legacy";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Operations";
                table.Rows[2].Cells[0].Paragraphs[0].Text = "Retention";
                table.Rows[2].Cells[1].Paragraphs[0].Text = "Legal";
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_row_deleted_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                WordTable table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Retention";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Legal";
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_row_deleted_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableRow &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted &&
                finding.SourceText == "Legacy | Operations");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            TableRow deletedRow = body.Descendants<Table>().First().Elements<TableRow>().ElementAt(1);
            Assert.Contains(deletedRow.Elements<TableCell>().ElementAt(0).Descendants<DeletedRun>(), run => run.InnerText == "Legacy" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(deletedRow.Elements<TableCell>().ElementAt(1).Descendants<DeletedRun>(), run => run.InnerText == "Operations" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains("Retention", body.Descendants<Table>().First().Elements<TableRow>().ElementAt(2).InnerText, StringComparison.Ordinal);

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForInsertedTables() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_inserted_whole_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Controls");
                WordTable table = document.AddTable(1, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_inserted_whole_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Controls");
                WordTable table = document.AddTable(1, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                WordTable inserted = document.AddTable(2, 2);
                inserted.Rows[0].Cells[0].Paragraphs[0].Text = "Escalation";
                inserted.Rows[0].Cells[1].Paragraphs[0].Text = "Support";
                inserted.Rows[1].Cells[0].Paragraphs[0].Text = "Review";
                inserted.Rows[1].Cells[1].Paragraphs[0].Text = "Quarterly";
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_inserted_whole_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Table &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText != null &&
                finding.TargetText.Contains("Escalation", StringComparison.Ordinal));

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Table insertedTable = body.Descendants<Table>().ElementAt(1);
            Assert.Contains(insertedTable.Descendants<InsertedRun>(), run => run.InnerText == "Escalation" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(insertedTable.Descendants<InsertedRun>(), run => run.InnerText == "Quarterly" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(insertedTable.Descendants<DeletedRun>(), run => run.InnerText.Contains("Escalation", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForDeletedTables() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_deleted_whole_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Controls");
                WordTable table = document.AddTable(1, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                WordTable deleted = document.AddTable(2, 2);
                deleted.Rows[0].Cells[0].Paragraphs[0].Text = "Legacy";
                deleted.Rows[0].Cells[1].Paragraphs[0].Text = "Operations";
                deleted.Rows[1].Cells[0].Paragraphs[0].Text = "Archive";
                deleted.Rows[1].Cells[1].Paragraphs[0].Text = "Annual";
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_deleted_whole_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Controls");
                WordTable table = document.AddTable(1, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_table_deleted_whole_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Table &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted &&
                finding.SourceText != null &&
                finding.SourceText.Contains("Legacy", StringComparison.Ordinal));

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Assert.Equal(2, body.Descendants<Table>().Count());
            Table deletedTable = body.Descendants<Table>().ElementAt(1);
            Assert.Contains(deletedTable.Descendants<DeletedRun>(), run => run.InnerText == "Legacy" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(deletedTable.Descendants<DeletedRun>(), run => run.InnerText == "Annual" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(deletedTable.Descendants<InsertedRun>(), run => run.InnerText.Contains("Legacy", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForInsertedNestedTables() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_nested_table_inserted_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                WordTable outer = document.AddTable(1, 1);
                outer.Rows[0].Cells[0].Paragraphs[0].Text = "Nested controls";
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_nested_table_inserted_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                WordTable outer = document.AddTable(1, 1);
                outer.Rows[0].Cells[0].Paragraphs[0].Text = "Nested controls";
                WordTable nested = outer.Rows[0].Cells[0].AddTable(2, 2);
                nested.Rows[0].Cells[0].Paragraphs[0].Text = "Escalation";
                nested.Rows[0].Cells[1].Paragraphs[0].Text = "Support";
                nested.Rows[1].Cells[0].Paragraphs[0].Text = "Review";
                nested.Rows[1].Cells[1].Paragraphs[0].Text = "Quarterly";
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_nested_table_inserted_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Table &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.Location == "table[1]" &&
                finding.TargetText != null &&
                finding.TargetText.Contains("Escalation", StringComparison.Ordinal));

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Table outerTable = Assert.Single(body.Elements<Table>());
            TableCell outerCell = Assert.Single(outerTable.Elements<TableRow>().Single().Elements<TableCell>());
            Table insertedNestedTable = Assert.Single(outerCell.Elements<Table>());
            Assert.Contains(outerCell.Elements<Paragraph>(), paragraph => paragraph.InnerText == "Nested controls");
            Assert.Contains(insertedNestedTable.Descendants<InsertedRun>(), run => run.InnerText == "Escalation" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(insertedNestedTable.Descendants<InsertedRun>(), run => run.InnerText == "Quarterly" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(insertedNestedTable.Descendants<DeletedRun>(), run => run.InnerText.Contains("Escalation", StringComparison.Ordinal));
            Assert.DoesNotContain(body.Elements<Table>().Skip(1), table => table.InnerText.Contains("Escalation", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForModifiedNestedTableCells() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_nested_table_cell_modified_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                WordTable outer = document.AddTable(1, 1);
                outer.Rows[0].Cells[0].Paragraphs[0].Text = "Nested controls";
                WordTable nested = outer.Rows[0].Cells[0].AddTable(2, 2);
                nested.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                nested.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                nested.Rows[1].Cells[0].Paragraphs[0].Text = "Retention";
                nested.Rows[1].Cells[1].Paragraphs[0].Text = "Legal";
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_nested_table_cell_modified_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                WordTable outer = document.AddTable(1, 1);
                outer.Rows[0].Cells[0].Paragraphs[0].Text = "Nested controls";
                WordTable nested = outer.Rows[0].Cells[0].AddTable(2, 2);
                nested.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                nested.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
                nested.Rows[1].Cells[0].Paragraphs[0].Text = "Retention";
                nested.Rows[1].Cells[1].Paragraphs[0].Text = "Compliance";
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_nested_table_cell_modified_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.TableCell &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.Location == "table[1]/row[1]/cell[1]" &&
                finding.SourceText == "Legal" &&
                finding.TargetText == "Compliance");

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Table outerTable = Assert.Single(body.Elements<Table>());
            TableCell outerCell = Assert.Single(outerTable.Elements<TableRow>().Single().Elements<TableCell>());
            Table nestedTable = Assert.Single(outerCell.Elements<Table>());
            TableCell changedNestedCell = nestedTable.Elements<TableRow>().ElementAt(1).Elements<TableCell>().ElementAt(1);
            Assert.Contains(outerCell.Elements<Paragraph>(), paragraph => paragraph.InnerText == "Nested controls");
            Assert.Contains(changedNestedCell.Descendants<DeletedRun>(), run => run.InnerText == "Legal" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(changedNestedCell.Descendants<InsertedRun>(), run => run.InnerText == "Compliance" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(body.Elements<Table>().Skip(1), table => table.InnerText.Contains("Compliance", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void CompareStructureCreatesInPlaceTargetRedlineForDeletedNestedTables() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_nested_table_deleted_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                WordTable outer = document.AddTable(1, 1);
                outer.Rows[0].Cells[0].Paragraphs[0].Text = "Nested controls";
                WordTable nested = outer.Rows[0].Cells[0].AddTable(2, 2);
                nested.Rows[0].Cells[0].Paragraphs[0].Text = "Legacy";
                nested.Rows[0].Cells[1].Paragraphs[0].Text = "Operations";
                nested.Rows[1].Cells[0].Paragraphs[0].Text = "Archive";
                nested.Rows[1].Cells[1].Paragraphs[0].Text = "Annual";
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_nested_table_deleted_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                WordTable outer = document.AddTable(1, 1);
                outer.Rows[0].Cells[0].Paragraphs[0].Text = "Nested controls";
                document.Save(false);
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_inplace_nested_table_deleted_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests"
                });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Table &&
                finding.ChangeKind == WordComparisonChangeKind.Deleted &&
                finding.Location == "table[1]" &&
                finding.SourceText != null &&
                finding.SourceText.Contains("Legacy", StringComparison.Ordinal));

            using WordDocument redline = WordDocument.Load(outputPath, readOnly: true);
            Body body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            Table outerTable = Assert.Single(body.Elements<Table>());
            TableCell outerCell = Assert.Single(outerTable.Elements<TableRow>().Single().Elements<TableCell>());
            Table deletedNestedTable = Assert.Single(outerCell.Elements<Table>());
            Assert.Contains(outerCell.Elements<Paragraph>(), paragraph => paragraph.InnerText == "Nested controls");
            Assert.Contains(deletedNestedTable.Descendants<DeletedRun>(), run => run.InnerText == "Legacy" && run.Author?.Value == "OfficeIMO Tests");
            Assert.Contains(deletedNestedTable.Descendants<DeletedRun>(), run => run.InnerText == "Annual" && run.Author?.Value == "OfficeIMO Tests");
            Assert.DoesNotContain(deletedNestedTable.Descendants<InsertedRun>(), run => run.InnerText.Contains("Legacy", StringComparison.Ordinal));
            Assert.DoesNotContain(body.Elements<Table>().Skip(1), table => table.InnerText.Contains("Legacy", StringComparison.Ordinal));

            var errors = redline.ValidateDocument();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        private static void WrapCellInCellContentControl(WordTableCell cell, string alias) {
            TableCell wrappedCell = (TableCell)cell._tableCell.CloneNode(true);
            var contentControl = new SdtCell(
                new SdtProperties(new SdtAlias { Val = alias }),
                new SdtContentCell(wrappedCell));

            cell._tableCell.InsertBeforeSelf(contentControl);
            cell._tableCell.Remove();
        }

        private static void WrapRowInRowContentControl(WordTableRow row, string alias) {
            TableRow wrappedRow = (TableRow)row._tableRow.CloneNode(true);
            var contentControl = new SdtRow(
                new SdtProperties(new SdtAlias { Val = alias }),
                new SdtContentRow(wrappedRow));

            row._tableRow.InsertBeforeSelf(contentControl);
            row._tableRow.Remove();
        }

        private static void AddNestedRunContentControl(WordDocument document, string value) {
            var innerControl = new SdtRun(
                new SdtProperties(
                    new SdtAlias { Val = "Nested value" },
                    new Tag { Val = "NestedValue" }),
                new SdtContentRun(new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve })));

            var outerControl = new SdtRun(
                new SdtProperties(
                    new SdtAlias { Val = "Nested container" },
                    new Tag { Val = "NestedContainer" }),
                new SdtContentRun(
                    new Run(new Text("Client: ") { Space = SpaceProcessingModeValues.Preserve }),
                    innerControl));

            document._document.Body!.Append(new Paragraph(outerControl));
        }

        private static void AddTextBoxRunContentControl(WordDocument document, string value) {
            document.AddTextBox(value);
            TextBoxContent textBoxContent = document._document.Body!.Descendants<TextBoxContent>().Last();
            Paragraph paragraph = textBoxContent.Elements<Paragraph>().First();
            paragraph.RemoveAllChildren<Run>();
            paragraph.Append(new SdtRun(
                new SdtProperties(
                    new SdtAlias { Val = "Text box status" },
                    new Tag { Val = "TextBoxStatus" }),
                new SdtContentRun(new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve }))));
        }

        private static void AddTextBoxBlockContentControl(WordDocument document, string value) {
            document.AddTextBox(value);
            TextBoxContent textBoxContent = document._document.Body!.Descendants<TextBoxContent>().Last();
            textBoxContent.RemoveAllChildren();
            textBoxContent.Append(new SdtBlock(
                new SdtProperties(
                    new SdtAlias { Val = "Text box review status" },
                    new Tag { Val = "TextBoxReviewStatus" }),
                new SdtContentBlock(
                    new Paragraph(
                        new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve })))));
        }

        private static void ReplaceLastFootnoteWithTable(WordDocument document, params string[][] rows) {
            Footnote footnote = document._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!
                .Elements<Footnote>()
                .Where(item => item.Type == null || item.Type.Value == FootnoteEndnoteValues.Normal)
                .Last();
            footnote.RemoveAllChildren();
            footnote.Append(new Paragraph(new Run(new FootnoteReferenceMark())));
            footnote.Append(CreateRawTable(rows));
        }

        private static void ReplaceLastEndnoteWithTable(WordDocument document, params string[][] rows) {
            Endnote endnote = document._wordprocessingDocument.MainDocumentPart!.EndnotesPart!.Endnotes!
                .Elements<Endnote>()
                .Where(item => item.Type == null || item.Type.Value == FootnoteEndnoteValues.Normal)
                .Last();
            endnote.RemoveAllChildren();
            endnote.Append(new Paragraph(new Run(new EndnoteReferenceMark())));
            endnote.Append(CreateRawTable(rows));
        }

        private static Table CreateRawTable(params string[][] rows) {
            int columnCount = rows.Length == 0 ? 1 : rows.Max(row => row.Length);
            var table = new Table(
                new TableProperties(new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto }),
                new TableGrid(Enumerable.Range(0, columnCount).Select(_ => new GridColumn { Width = "2400" })));

            foreach (string[] row in rows) {
                var tableRow = new TableRow();
                foreach (string value in row) {
                    tableRow.Append(new TableCell(
                        new TableCellProperties(new TableCellWidth { Width = "2400", Type = TableWidthUnitValues.Dxa }),
                        new Paragraph(new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve }))));
                }

                table.Append(tableRow);
            }

            return table;
        }

        private static DocumentFormat.OpenXml.Wordprocessing.Drawing CreateNonImageDrawing() {
            return new DocumentFormat.OpenXml.Wordprocessing.Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = 914400L, Cy = 457200L },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties { Id = 9001U, Name = "Non-image drawing" },
                    new DW.NonVisualGraphicFrameDrawingProperties(),
                    new A.Graphic(
                        new A.GraphicData(
                            new A.ShapeProperties()) {
                            Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
                        })));
        }
    }
}
