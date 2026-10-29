using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_LockedTocPreservesNestedFieldsButUpdatesIndependentTitle(bool regenerate) {
            using WordDocument document = WordDocument.Create();
            document.AddTableOfContent();
            document.AddParagraph("Heading").Style = WordParagraphStyles.Heading1;
            document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true });
            SdtBlock block = document.TableOfContent!.SdtBlock;
            block.Descendants<FieldChar>().First(x => x.FieldCharType?.Value == FieldCharValues.Begin).FieldLock = true;
            var titleField = new SimpleField(new Run(new Text("old"))) { Instruction = "DATE" };
            block.GetFirstChild<SdtContentBlock>()!.PrependChild(new Paragraph(titleField));
            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions {
                UpdateTablesOfContents = regenerate, CurrentDateTime = new DateTime(2020, 2, 3)
            });
            Assert.Equal(WordFieldUpdateStatus.Updated, report.Results.Single(x => x.FieldType == WordFieldType.Date).Status);
            Assert.Equal(WordFieldUpdateStatus.Skipped, report.Results.Single(x => x.FieldType == WordFieldType.TOC).Status);
            var pageReferences = report.Results.Where(x => x.FieldType == WordFieldType.PageRef).ToList();
            Assert.NotEmpty(pageReferences);
            Assert.All(pageReferences, x => Assert.Equal(WordFieldUpdateStatus.Skipped, x.Status));
            Assert.NotEqual("old", titleField.InnerText);
        }

        [Fact]
        public void Test_UpdateFieldsLayoutUsesRestartedSectionPageNumbers() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("First section");
            var section = document.AddSection();
            section.AddPageNumbering(10);
            section.AddParagraph("Target").AddBookmark("RestartTarget");
            section.AddParagraph("Reference").AddField(WordFieldType.PageRef, parameters: new List<string> { "RestartTarget" });
            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions { UseLayoutPagination = true });
            Assert.Equal("10", report.Results.Single(x => x.FieldType == WordFieldType.PageRef).ResultText);
        }

        [Fact]
        public void Test_UpdateFieldsLayoutUsesSectionNumberFormat() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("First section");
            var section = document.AddSection();
            section.AddPageNumbering(10, WordNumberFormat.LowerRoman);
            section.AddParagraph("Target").AddBookmark("RomanTarget");
            section.AddParagraph("Reference").AddField(WordFieldType.PageRef, parameters: new List<string> { "RomanTarget" });
            section.AddParagraph("Arabic override").AddField(WordFieldType.PageRef, WordFieldFormat.Arabic, parameters: new List<string> { "RomanTarget" });
            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions { UseLayoutPagination = true });
            var references = report.Results.Where(x => x.FieldType == WordFieldType.PageRef).ToList();
            Assert.Equal(new[] { "x", "10" }, references.Select(x => x.ResultText));
        }

        [Theory]
        [InlineData(WordNumberFormat.Ordinal)]
        [InlineData(WordNumberFormat.DecimalZero)]
        public void Test_UpdateFieldsLayoutLeavesUnsupportedSectionNumberFormatUnchanged(WordNumberFormat format) {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("First section");
            var section = document.AddSection();
            section.AddPageNumbering(10, format);
            section.AddParagraph("Target").AddBookmark("OrdinalTarget");
            section.AddParagraph("Reference").AddField(WordFieldType.PageRef, parameters: new List<string> { "OrdinalTarget" });
            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions { UseLayoutPagination = true });
            Assert.Equal(WordFieldUpdateStatus.Unsupported, report.Results.Single(x => x.FieldType == WordFieldType.PageRef).Status);
        }

        [Fact]
        public void Test_LockedTocKeepsNestedFieldInSameRunAsEnd() {
            using WordDocument document = WordDocument.Create();
            Paragraph paragraph = document.AddParagraph()._paragraph;
            paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin, FieldLock = true },
                new FieldCode(" TOC "), new FieldChar { FieldCharType = FieldCharValues.Separate }));
            paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }, new FieldCode(" DATE "),
                new FieldChar { FieldCharType = FieldCharValues.Separate }, new Text("old"),
                new FieldChar { FieldCharType = FieldCharValues.End }, new FieldChar { FieldCharType = FieldCharValues.End }));
            string before = paragraph.OuterXml;
            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions { CurrentDateTime = new DateTime(2020, 2, 3) });
            Assert.Equal(WordFieldUpdateStatus.Skipped, report.Results.Single(x => x.FieldType == WordFieldType.Date).Status);
            Assert.Equal(before, paragraph.OuterXml);
        }

        [Fact]
        public void Test_UpdateFieldsLayoutDoesNotNormalizeOriginalTables() {
            using WordDocument document = WordDocument.Create();
            document.AddTable(1, 2);
            Table table = document._document.Body!.Descendants<Table>().First();
            table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().First().Width = "3500";
            string before = table.OuterXml;
            document.UpdateFields(new WordFieldUpdateOptions { UseLayoutPagination = true });
            Assert.Equal(before, table.OuterXml);
        }

        [Fact]
        public void Test_UpdateFieldsDoesNotRegenerateLockedToc() {
            using WordDocument document = WordDocument.Create();
            document.AddTableOfContent();
            document.AddParagraph("Heading").Style = WordParagraphStyles.Heading1;
            document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true });
            SdtBlock block = document.TableOfContent!.SdtBlock;
            FieldChar begin = block.Descendants<FieldChar>().First(x => x.FieldCharType?.Value == FieldCharValues.Begin);
            begin.FieldLock = true;
            string before = block.OuterXml;
            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions { UpdateTablesOfContents = true });
            Assert.Equal(before, block.OuterXml);
            Assert.Equal(WordFieldUpdateStatus.Skipped, report.Results.Single(x => x.FieldType == WordFieldType.TOC).Status);
        }

        [Fact]
        public void Test_UpdateFieldsRegeneratesTableOfContentsFromLayout() {
            string filePath = Path.Combine(_directoryWithFiles, "UpdateFieldsLayoutToc.docx");
            string? pageReference;
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddTableOfContent();
                document.AddPageBreak();
                document.AddParagraph("Layout Overview").Style = WordParagraphStyles.Heading1;
                for (int index = 0; index < 120; index++) {
                    document.AddParagraph("Body line " + index + " that fills the page so the next heading flows onto a later page without an explicit break.");
                }

                document.AddParagraph("Layout Details").Style = WordParagraphStyles.Heading2;
                document.AddPageBreak();
                document.AddParagraph("Layout Appendix").Style = WordParagraphStyles.Heading1;
                document.AddParagraph("Appendix target").AddBookmark("AppendixTarget");
                document.AddParagraph("See page ").AddField(WordFieldType.PageRef, parameters: new List<string> { "AppendixTarget" });

                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions {
                    UpdateTablesOfContents = true,
                    UseLayoutPagination = true
                });

                WordFieldUpdateResult toc = report.Results.Single(_ => _.FieldType == WordFieldType.TOC);
                Assert.Equal(WordFieldUpdateStatus.Updated, toc.Status);
                Assert.All(report.Results.Where(_ => _.FieldType == WordFieldType.PageRef), _ => {
                    Assert.Equal(WordFieldUpdateStatus.Updated, _.Status);
                    Assert.Equal(WordFieldEvaluationBasis.LayoutPagination, _.EvaluationBasis);
                });
                pageReference = report.Results.Single(_ => _.FieldType == WordFieldType.PageRef && _.InstructionText.Contains("AppendixTarget")).ResultText;
                Assert.False(document.Settings.UpdateFieldsOnOpen);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                SdtBlock block = document.TableOfContent!.SdtBlock;
                Dictionary<string, int> pages = ReadTocEntryPages(block);

                Assert.Equal(new[] { "Layout Overview", "Layout Details", "Layout Appendix" }, pages.Keys.ToArray());
                Assert.Equal(2, pages["Layout Overview"]);
                // Only layout pagination can place this heading after page 2: there is no explicit break before it.
                Assert.True(pages["Layout Details"] > pages["Layout Overview"], string.Join(", ", pages));
                Assert.Equal(pages["Layout Details"] + 1, pages["Layout Appendix"]);

                Assert.Contains(block.Descendants<FieldCode>(), _ => _.Text.TrimStart().StartsWith("TOC", StringComparison.Ordinal));
                Assert.DoesNotContain(block.Descendants<FieldChar>(), _ => _.Dirty?.Value == true);
                Assert.Equal(pages["Layout Appendix"].ToString(), pageReference);

                var errors = document.ValidateDocument();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));

                // Updating again rebuilds the same entries from the regenerated field.
                document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true, UseLayoutPagination = true });
                Assert.Equal(pages, ReadTocEntryPages(document.TableOfContent!.SdtBlock));
            }
        }

        [Fact]
        public void Test_UpdateFieldsFillsWordGeneratedTableOfContentsFixture() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-outline-toc.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated TOC fixture: {sourcePath}");
            string filePath = Path.Combine(_directoryWithFiles, "UpdateFieldsLayoutTocFixture.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true, UseLayoutPagination = true });
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Dictionary<string, int> pages = ReadTocEntryPages(document.TableOfContent!.SdtBlock);
                Assert.NotEmpty(pages);
                Assert.All(pages.Values, _ => Assert.True(_ >= 1));
                Assert.Equal(pages.Values.OrderBy(_ => _).ToArray(), pages.Values.ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsKeepsCallerUpdateFieldsOnOpen() {
            using WordDocument document = WordDocument.Create();
            document.Settings.UpdateFieldsOnOpen = true;
            document.AddTableOfContent();
            document.AddParagraph("Kept heading").Style = WordParagraphStyles.Heading1;

            document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true });

            Assert.True(document.Settings.UpdateFieldsOnOpen);
        }

        [Fact]
        public void Test_UpdateFieldsKeepsUpdateFieldsOnOpenSetAfterTableOfContents() {
            using WordDocument document = WordDocument.Create();
            document.AddTableOfContent();
            document.AddParagraph("Kept heading").Style = WordParagraphStyles.Heading1;
            document.Settings.UpdateFieldsOnOpen = true;

            document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true });

            Assert.True(document.Settings.UpdateFieldsOnOpen);
        }

        [Fact]
        public void Test_UpdateFieldsClearsTableOfContentsUpdateOnOpenAfterSetLevels() {
            using WordDocument document = WordDocument.Create();
            WordTableOfContent toc = document.AddTableOfContent();
            toc.SetLevels(1, 2);
            document.AddParagraph("Heading").Style = WordParagraphStyles.Heading1;

            document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true });

            Assert.False(document.Settings.UpdateFieldsOnOpen);
        }

        [Fact]
        public void Test_UpdateFieldsPlacesTableCellBookmarksOnlyWhenLayoutPinsThePage() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Start").AddBookmark("StartTarget");
            WordTable shortTable = document.AddTable(2, 2);
            shortTable.Rows[1].Cells[0].Paragraphs[0].AddText("Short cell").AddBookmark("ShortCellTarget");
            document.AddParagraph("Middle").AddBookmark("MiddleTarget");
            WordTable longTable = document.AddTable(150, 1);
            longTable.Rows[149].Cells[0].Paragraphs[0].AddText("Long cell").AddBookmark("LongCellTarget");
            document.AddPageBreak();
            document.AddParagraph("End").AddBookmark("EndTarget");
            document.AddParagraph("Short ").AddField(WordFieldType.PageRef, parameters: new List<string> { "ShortCellTarget" });
            document.AddParagraph("Long ").AddField(WordFieldType.PageRef, parameters: new List<string> { "LongCellTarget" });

            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions { UseLayoutPagination = true });

            WordFieldUpdateResult shortCell = report.Results.Single(_ => _.InstructionText.Contains("ShortCellTarget"));
            Assert.Equal(WordFieldUpdateStatus.Updated, shortCell.Status);
            Assert.Equal("1", shortCell.ResultText);
            Assert.Equal(WordFieldEvaluationBasis.LayoutPagination, shortCell.EvaluationBasis);

            WordFieldUpdateResult longCell = report.Results.Single(_ => _.InstructionText.Contains("LongCellTarget"));
            // The bookmarks around the long table are on different pages, so its cell bookmark has no reliable page.
            Assert.Equal(WordFieldUpdateStatus.Skipped, longCell.Status);
        }

        private static Dictionary<string, int> ReadTocEntryPages(SdtBlock block) {
            var pages = new Dictionary<string, int>();
            foreach (Hyperlink hyperlink in block.Descendants<Hyperlink>().Where(_ => !string.IsNullOrEmpty(_.Anchor?.Value))) {
                Assert.Contains(hyperlink.Descendants<FieldCode>(), _ => _.Text.Contains("PAGEREF " + hyperlink.Anchor!.Value, StringComparison.Ordinal));
                var texts = hyperlink.Descendants<Text>().Select(_ => _.Text).ToList();
                pages[texts.First()] = int.Parse(texts.Last());
            }

            return pages;
        }
    }
}
