using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
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
