using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CompareStructureOptionsIgnoreWhitespaceAndCaseTextDifferences() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_options_text_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Policy   Owner");
                document.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_options_text_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("policy owner");
                document.Save();
            }

            WordComparisonResult defaultResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath);
            Assert.Contains(defaultResult.Findings, finding => finding.Scope == WordComparisonScope.Paragraph);

            WordComparisonResult ignoredResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                IgnoreWhitespace = true,
                IgnoreCase = true
            });

            Assert.False(ignoredResult.HasChanges);
        }

        [Fact]
        public void CompareStructureOptionsCanDisableFieldAndContentControlFindings() {
            string fieldSourcePath = Path.Combine(_directoryWithFiles, "compare_structure_options_field_source.docx");
            CreateDocumentWithSimpleField(fieldSourcePath, " AUTHOR ", "Alice");
            string fieldTargetPath = Path.Combine(_directoryWithFiles, "compare_structure_options_field_target.docx");
            CreateDocumentWithSimpleField(fieldTargetPath, " TITLE ", "Quarterly report");

            WordComparisonResult fieldResult = WordDocumentComparer.CompareStructure(fieldSourcePath, fieldTargetPath, new WordComparisonOptions {
                CompareFields = false
            });
            Assert.DoesNotContain(fieldResult.Findings, finding => finding.Scope == WordComparisonScope.Field);

            string controlSourcePath = Path.Combine(_directoryWithFiles, "compare_structure_options_control_source.docx");
            CreateDocumentWithBoundContentControl(
                controlSourcePath,
                alias: "Client",
                tag: "Client.Name",
                storeItemId: "{11111111-1111-1111-1111-111111111111}",
                xpath: "/root/client/name",
                text: "Contoso");
            string controlTargetPath = Path.Combine(_directoryWithFiles, "compare_structure_options_control_target.docx");
            CreateDocumentWithBoundContentControl(
                controlTargetPath,
                alias: "Customer",
                tag: "Customer.Name",
                storeItemId: "{22222222-2222-2222-2222-222222222222}",
                xpath: "/root/customer/name",
                text: "Contoso");

            WordComparisonResult controlResult = WordDocumentComparer.CompareStructure(controlSourcePath, controlTargetPath, new WordComparisonOptions {
                CompareContentControls = false
            });
            Assert.DoesNotContain(controlResult.Findings, finding => finding.Scope == WordComparisonScope.ContentControl);
        }

        [Fact]
        public void CompareStructureOptionsCanDisableRunFormattingFindings() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_options_format_source.docx");
            CreateDocumentWithSingleRun(sourcePath, bold: false);
            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_options_format_target.docx");
            CreateDocumentWithSingleRun(targetPath, bold: true);

            WordComparisonResult defaultResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath);
            Assert.Contains(defaultResult.Findings, finding =>
                finding.Scope == WordComparisonScope.Run &&
                finding.Message == "Run formatting changed.");

            WordComparisonResult ignoredResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                CompareRunFormatting = false
            });
            Assert.DoesNotContain(ignoredResult.Findings, finding => finding.Scope == WordComparisonScope.Run);
        }

        [Fact]
        public void CompareStructureOptionsCanDisableParagraphStyleIdFindings() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_options_paragraph_style_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Same styled text").Style = WordParagraphStyles.Normal;
                document.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_options_paragraph_style_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Same styled text").Style = WordParagraphStyles.Heading1;
                document.Save();
            }

            WordComparisonResult defaultResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath);
            Assert.Contains(defaultResult.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.Message == "Paragraph style id changed.");

            WordComparisonResult ignoredResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                CompareParagraphStyleIds = false
            });
            Assert.DoesNotContain(ignoredResult.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.Message == "Paragraph style id changed.");
        }

        [Fact]
        public void CompareStructureOptionsCanDisableRunStyleIdFindings() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_options_run_style_source.docx");
            CreateDocumentWithRunStyle(sourcePath, runStyleId: "Emphasis");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_options_run_style_target.docx");
            CreateDocumentWithRunStyle(targetPath, runStyleId: "Strong");

            WordComparisonResult defaultResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath);
            Assert.Contains(defaultResult.Findings, finding =>
                finding.Scope == WordComparisonScope.Run &&
                finding.Message == "Run formatting changed.");

            WordComparisonResult ignoredResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                CompareRunStyleIds = false
            });
            Assert.DoesNotContain(ignoredResult.Findings, finding =>
                finding.Scope == WordComparisonScope.Run &&
                finding.Message == "Run formatting changed.");
        }

        [Fact]
        public void CompareStructureReportsEffectiveFormattingChangesFromStyleDefinitions() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_effective_format_source.docx");
            CreateDocumentWithInheritedComparisonStyle(sourcePath, paragraphSpacingAfter: "120", runColor: "1F4E79");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_effective_format_target.docx");
            CreateDocumentWithInheritedComparisonStyle(targetPath, paragraphSpacingAfter: "240", runColor: "C00000");

            WordComparisonResult defaultResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(defaultResult.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.Message == "Paragraph effective formatting changed.");
            Assert.Contains(defaultResult.Findings, finding =>
                finding.Scope == WordComparisonScope.Run &&
                finding.Message == "Run formatting changed." &&
                finding.SourceText == "Inherited formatting text" &&
                finding.TargetText == "Inherited formatting text");

            WordComparisonResult directOnlyResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                CompareEffectiveFormatting = false
            });
            Assert.False(directOnlyResult.HasChanges);
        }

        [Fact]
        public void CompareStructureOptionsCanDisableImageFindings() {
            string logoPath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string replacementPath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_options_image_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph().AddImage(logoPath, 80, 40);
                document.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_options_image_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph().AddImage(replacementPath, 80, 40);
                document.Save();
            }

            WordComparisonResult defaultResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath);
            Assert.Contains(defaultResult.Findings, finding => finding.Scope == WordComparisonScope.Image);

            WordComparisonResult ignoredResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                CompareImages = false
            });
            Assert.DoesNotContain(ignoredResult.Findings, finding => finding.Scope == WordComparisonScope.Image);
        }

        [Fact]
        public void CompareStructureOptionsCanFilterOutputByIncludedAndExcludedScopes() {
            string logoPath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            string replacementPath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_options_scope_filter_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Source paragraph");
                document.AddParagraph().AddImage(logoPath, 80, 40);
                document.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_options_scope_filter_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Target paragraph");
                document.AddParagraph().AddImage(replacementPath, 80, 40);
                document.Save();
            }

            WordComparisonResult defaultResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath);
            Assert.Contains(defaultResult.Findings, finding => finding.Scope == WordComparisonScope.Paragraph);
            Assert.Contains(defaultResult.Findings, finding => finding.Scope == WordComparisonScope.Image);

            WordComparisonResult imageOnly = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                IncludedScopes = new HashSet<WordComparisonScope> {
                    WordComparisonScope.Image
                }
            });
            Assert.NotEmpty(imageOnly.Findings);
            Assert.All(imageOnly.Findings, finding => Assert.Equal(WordComparisonScope.Image, finding.Scope));

            WordComparisonResult withoutImages = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                ExcludedScopes = new HashSet<WordComparisonScope> {
                    WordComparisonScope.Image
                }
            });
            Assert.Contains(withoutImages.Findings, finding => finding.Scope == WordComparisonScope.Paragraph);
            Assert.DoesNotContain(withoutImages.Findings, finding => finding.Scope == WordComparisonScope.Image);
        }

        [Fact]
        public void CompareStructureOptionsCanIgnoreGeneratedIdsAndVolatileMetadata() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_options_generated_source.docx");
            CreateDocumentWithStableReviewContent(
                sourcePath,
                commentDate: new DateTime(2026, 1, 1, 10, 0, 0, DateTimeKind.Utc),
                commentParaId: "AAAA1111",
                revisionDate: new DateTime(2026, 1, 1, 11, 0, 0, DateTimeKind.Utc),
                revisionId: "101");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_options_generated_target.docx");
            CreateDocumentWithStableReviewContent(
                targetPath,
                commentDate: new DateTime(2026, 2, 1, 10, 0, 0, DateTimeKind.Utc),
                commentParaId: "BBBB2222",
                revisionDate: new DateTime(2026, 2, 1, 11, 0, 0, DateTimeKind.Utc),
                revisionId: "202");

            WordComparisonResult defaultResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath);
            Assert.Contains(defaultResult.Findings, finding => finding.Scope == WordComparisonScope.Comment);
            Assert.Contains(defaultResult.Findings, finding => finding.Scope == WordComparisonScope.Revision);

            WordComparisonResult ignoredResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                CompareGeneratedIds = false,
                CompareVolatileMetadata = false
            });

            Assert.DoesNotContain(ignoredResult.Findings, finding => finding.Scope == WordComparisonScope.Comment);
            Assert.DoesNotContain(ignoredResult.Findings, finding => finding.Scope == WordComparisonScope.Revision);
        }

        private static void CreateDocumentWithSingleRun(string path, bool bold) {
            using WordDocument document = WordDocument.Create(path);
            document.AddParagraph("Same text");
            document.Save();

            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, true);
            Run run = wordDocument.MainDocumentPart!.Document.Body!.Descendants<Run>().First();
            run.RunProperties ??= new RunProperties();
            run.RunProperties.Bold = bold ? new Bold() : null;
            wordDocument.MainDocumentPart.Document.Save();
        }

        private static void CreateDocumentWithRunStyle(string path, string runStyleId) {
            using (WordDocument document = WordDocument.Create(path)) {
                document.AddParagraph("Same run style text");
                document.Save();
            }

            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, true);
            Run run = wordDocument.MainDocumentPart!.Document.Body!.Descendants<Run>().First();
            run.RunProperties ??= new RunProperties();
            run.RunProperties.RunStyle = new RunStyle {
                Val = runStyleId
            };
            wordDocument.MainDocumentPart.Document.Save();
        }

        private static void CreateDocumentWithInheritedComparisonStyle(string path, string paragraphSpacingAfter, string runColor) {
            using (WordDocument document = WordDocument.Create(path)) {
                document.AddParagraph("Inherited formatting text");
                document.Save();
            }

            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
            StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
            Styles styles = stylePart.Styles ??= new Styles();

            foreach (Style existing in styles.Elements<Style>()
                .Where(style =>
                    string.Equals(style.StyleId?.Value, "OfficeIMOEffectiveBase", StringComparison.Ordinal) ||
                    string.Equals(style.StyleId?.Value, "OfficeIMOEffectiveDerived", StringComparison.Ordinal))
                .ToList()) {
                existing.Remove();
            }

            styles.Append(
                new Style(
                    new StyleName { Val = "OfficeIMO Effective Base" },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(
                        new SpacingBetweenLines { After = paragraphSpacingAfter }),
                    new StyleRunProperties(
                        new DocumentFormat.OpenXml.Wordprocessing.Color { Val = runColor })) {
                    Type = StyleValues.Paragraph,
                    StyleId = "OfficeIMOEffectiveBase",
                    CustomStyle = true
                });
            styles.Append(
                new Style(
                    new StyleName { Val = "OfficeIMO Effective Derived" },
                    new BasedOn { Val = "OfficeIMOEffectiveBase" }) {
                    Type = StyleValues.Paragraph,
                    StyleId = "OfficeIMOEffectiveDerived",
                    CustomStyle = true
                });
            stylePart.Styles.Save();

            Paragraph paragraph = mainPart.Document.Body!.Elements<Paragraph>().First();
            paragraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId { Val = "OfficeIMOEffectiveDerived" });
            mainPart.Document.Save();
        }

        private static void CreateDocumentWithStableReviewContent(
            string path,
            DateTime commentDate,
            string commentParaId,
            DateTime revisionDate,
            string revisionId) {
            using (WordDocument document = WordDocument.Create(path)) {
                document.AddParagraph("Review target").AddComment("Alice Reviewer", "AR", "Same review note.");
                document.AddParagraph("Tracked ").AddInsertedText("Same revision", "Alice Reviewer", revisionDate);
                document.Save();
            }

            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
            Comment comment = mainPart.WordprocessingCommentsPart!.Comments!.Elements<Comment>().Single();
            comment.Date = commentDate;
            mainPart.WordprocessingCommentsPart.Comments.Save();

            CommentEx commentEx = mainPart.WordprocessingCommentsExPart!.CommentsEx!.Elements<CommentEx>().Single();
            commentEx.ParaId = commentParaId;
            mainPart.WordprocessingCommentsExPart.CommentsEx.Save();

            InsertedRun inserted = mainPart.Document.Body!.Descendants<InsertedRun>().Single();
            inserted.Id = revisionId;
            inserted.Date = revisionDate;
            mainPart.Document.Save();
        }
    }
}
