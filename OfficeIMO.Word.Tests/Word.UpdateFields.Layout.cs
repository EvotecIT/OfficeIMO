using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
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
