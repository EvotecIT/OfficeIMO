using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_TableOfContent_RefreshEntriesGeneratesHeadingEntriesAndBookmarks() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntries.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 2);
                toc.Text = "Contents";

                document.AddParagraph("Overview").SetStyle(WordParagraphStyles.Heading1);
                document.AddPageBreak();
                document.AddParagraph("Details").SetStyle(WordParagraphStyles.Heading2);
                document.AddParagraph("Appendix").SetStyle(WordParagraphStyles.Heading3);

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(1, report.SkippedHeadingCount);
                Assert.Equal(new[] { "Overview", "Details" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.All(report.Entries, entry => Assert.StartsWith("_OfficeIMO_Toc_", entry.BookmarkName));

                AssertGeneratedEntries(toc, "Overview", "Details");
                Assert.DoesNotContain("Appendix", TocText(toc));
                Assert.All(report.Entries, entry => Assert.Contains(document._document.Body!.Descendants<BookmarkStart>(), bookmark => bookmark.Name == entry.BookmarkName));
                Assert.True(toc.SdtBlock.Descendants<SimpleField>().Single().Dirty?.Value);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Equal("Contents", toc.Text);
                AssertGeneratedEntries(toc, "Overview", "Details");
                Assert.DoesNotContain("Appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesCountsEveryExplicitPageBreak() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesMultipleBreaks.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 1);
                document.AddParagraph("First").SetStyle(WordParagraphStyles.Heading1);
                AppendBodyParagraph(document, new Paragraph(
                    new Run(new Break { Type = BreakValues.Page }),
                    new Run(new Break { Type = BreakValues.Page })));
                document.AddParagraph("Second").SetStyle(WordParagraphStyles.Heading1);

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(new[] { "First", "Second" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 3 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesReplacesExistingEntriesAndHonorsLevelRange() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesReplace.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 3);

                document.AddParagraph("Initial").SetStyle(WordParagraphStyles.Heading1);
                Assert.Equal(1, toc.RefreshEntries().EntryCount);

                document.AddParagraph("Second").SetStyle(WordParagraphStyles.Heading2);
                WordTableOfContentRefreshReport expanded = toc.RefreshEntries();

                Assert.Equal(2, expanded.EntryCount);
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));

                toc.SetLevels(2, 2);
                WordTableOfContentRefreshReport filtered = toc.RefreshEntries();

                Assert.Single(filtered.Entries);
                Assert.Equal("Second", filtered.Entries[0].Text);
                Assert.DoesNotContain("Initial", TocText(toc));
                Assert.Contains("Second", TocText(toc));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.DoesNotContain("Initial", TocText(toc));
                Assert.Contains("Second", TocText(toc));
                Assert.Single(toc.SdtBlock.Descendants<Hyperlink>(), hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesUsesDirectOutlineLevels() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesOutlineLevels.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 2);

                AddOutlineLevelParagraph(document, "Outline root", 0);
                document.AddPageBreak();
                AddOutlineLevelParagraph(document, "Outline child", 1);
                AddOutlineLevelParagraph(document, "Outline appendix", 3);
                document.AddParagraph("Plain body text");

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(1, report.SkippedHeadingCount);
                Assert.Equal(new[] { "Outline root", "Outline child" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                AssertGeneratedEntries(toc, "Outline root", "Outline child");
                Assert.DoesNotContain("Outline appendix", TocText(toc));
                Assert.DoesNotContain("Plain body text", TocText(toc));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertGeneratedEntries(toc, "Outline root", "Outline child");
                Assert.DoesNotContain("Outline appendix", TocText(toc));
                Assert.DoesNotContain("Plain body text", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesEstimatesSectionBreakPageNumbers() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesSectionBreaks.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 1);

                document.AddParagraph("First Section").SetStyle(WordParagraphStyles.Heading1);
                AddSectionBreakParagraph(document, SectionMarkValues.NextPage);
                document.AddParagraph("Second Section").SetStyle(WordParagraphStyles.Heading1);
                AddSectionBreakParagraph(document, SectionMarkValues.OddPage);
                document.AddParagraph("Third Section").SetStyle(WordParagraphStyles.Heading1);
                AddSectionBreakParagraph(document, SectionMarkValues.Continuous);
                document.AddParagraph("Continuous Section").SetStyle(WordParagraphStyles.Heading1);

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(new[] { "First Section", "Second Section", "Third Section", "Continuous Section" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2, 3, 3 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertGeneratedEntries(toc, "First Section", "Second Section", "Third Section", "Continuous Section");
                Assert.Equal(4, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesSupportsWordGeneratedComplexToc() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-toc.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated TOC fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesWordGenerated.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Equal(1, toc.MinLevel);
                Assert.Equal(2, toc.MaxLevel);
                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("Word-generated TOC Overview", TocText(toc));
                Assert.Contains("Word-generated TOC Detail", TocText(toc));
                Assert.DoesNotContain("Word-generated TOC Appendix", TocText(toc));

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(1, report.SkippedHeadingCount);
                Assert.Equal(new[] { "Word-generated TOC Overview", "Word-generated TOC Detail" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                AssertGeneratedEntries(toc, "Word-generated TOC Overview", "Word-generated TOC Detail");
                Assert.DoesNotContain("Word-generated TOC Appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(toc.SdtBlock.Descendants<SimpleField>().Single().Dirty?.Value);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertGeneratedEntries(toc, "Word-generated TOC Overview", "Word-generated TOC Detail");
                Assert.DoesNotContain("Word-generated TOC Appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesSupportsSplitComplexTocInstruction() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshSplitComplexInstruction.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                AppendBodyParagraph(document, new Paragraph(
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" TO") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FieldCode("C \\o \"1-1\" \\h") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("No table of contents entries found.") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End })));
                document.AddParagraph("Split Instruction Heading").SetStyle(WordParagraphStyles.Heading1);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(1, report.EntryCount);
                Assert.Contains("Split Instruction Heading", TocText(toc));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesSupportsWordGeneratedTableCellHeadings() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-toc-table-cell.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated table-cell TOC fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesWordGeneratedTableCell.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Equal(1, toc.MinLevel);
                Assert.Equal(2, toc.MaxLevel);
                Assert.Contains(document._document.Body!.Descendants<Table>(), table =>
                    table.Descendants<Paragraph>().Any(paragraph =>
                        paragraph.Descendants<Text>().Any(text => text.Text.Contains("Word-generated TOC Detail", StringComparison.Ordinal))));
                Assert.Contains("Word-generated TOC Overview", TocText(toc));
                Assert.Contains("Word-generated TOC Detail", TocText(toc));
                Assert.DoesNotContain("Word-generated TOC Appendix", TocText(toc));

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(1, report.SkippedHeadingCount);
                Assert.Equal(new[] { "Word-generated TOC Overview", "Word-generated TOC Detail" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                WordTableOfContentEntry detailEntry = report.Entries.Single(entry => entry.Text == "Word-generated TOC Detail");
                Assert.StartsWith("_OfficeIMO_Toc_", detailEntry.BookmarkName);
                Assert.Contains(document._document.Body!.Descendants<Table>(), table =>
                    table.Descendants<BookmarkStart>().Any(bookmark => bookmark.Name == detailEntry.BookmarkName));
                AssertGeneratedEntries(toc, "Word-generated TOC Overview", "Word-generated TOC Detail");
                Assert.DoesNotContain("Word-generated TOC Appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(toc.SdtBlock.Descendants<SimpleField>().Single().Dirty?.Value);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertGeneratedEntries(toc, "Word-generated TOC Overview", "Word-generated TOC Detail");
                Assert.DoesNotContain("Word-generated TOC Appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesSupportsWordGeneratedContentControlHeadings() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-toc-content-control.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated content-control TOC fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesWordGeneratedContentControl.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Equal(1, toc.MinLevel);
                Assert.Equal(2, toc.MaxLevel);
                Assert.Contains(document._document.Body!.Descendants<SdtBlock>(), control =>
                    control.Descendants<Paragraph>().Any(paragraph =>
                        paragraph.ParagraphProperties?.ParagraphStyleId?.Val == "Heading2" &&
                        paragraph.Descendants<Text>().Any(text => text.Text.Contains("Word-generated TOC Detail", StringComparison.Ordinal))));
                Assert.Contains("Word-generated TOC Overview", TocText(toc));
                Assert.Contains("Word-generated TOC Detail", TocText(toc));
                Assert.DoesNotContain("Word-generated TOC Appendix", TocText(toc));

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(1, report.SkippedHeadingCount);
                Assert.Equal(new[] { "Word-generated TOC Overview", "Word-generated TOC Detail" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                WordTableOfContentEntry detailEntry = report.Entries.Single(entry => entry.Text == "Word-generated TOC Detail");
                Assert.StartsWith("_OfficeIMO_Toc_", detailEntry.BookmarkName);
                Assert.Contains(document._document.Body!.Descendants<SdtBlock>(), control =>
                    control.Descendants<BookmarkStart>().Any(bookmark => bookmark.Name == detailEntry.BookmarkName));
                AssertGeneratedEntries(toc, "Word-generated TOC Overview", "Word-generated TOC Detail");
                Assert.DoesNotContain("Word-generated TOC Appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(toc.SdtBlock.Descendants<SimpleField>().Single().Dirty?.Value);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertGeneratedEntries(toc, "Word-generated TOC Overview", "Word-generated TOC Detail");
                Assert.DoesNotContain("Word-generated TOC Appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesSupportsTextBoxHeadings() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesTextBoxHeadings.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 2);

                document.AddParagraph("Text-box TOC overview").SetStyle(WordParagraphStyles.Heading1);
                document.AddPageBreak();
                AppendBodyParagraph(document, CreateTextBoxHeadingParagraph("Text-box TOC detail", "Heading2"));
                document.AddParagraph("Text-box TOC appendix").SetStyle(WordParagraphStyles.Heading3);

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(1, report.SkippedHeadingCount);
                Assert.Equal(new[] { "Text-box TOC overview", "Text-box TOC detail" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                WordTableOfContentEntry detailEntry = report.Entries.Single(entry => entry.Text == "Text-box TOC detail");
                Assert.StartsWith("_OfficeIMO_Toc_", detailEntry.BookmarkName);
                Assert.Contains(document._document.Body!.Descendants<TextBoxContent>(), textBox =>
                    textBox.Descendants<BookmarkStart>().Any(bookmark => bookmark.Name == detailEntry.BookmarkName));
                AssertGeneratedEntries(toc, "Text-box TOC overview", "Text-box TOC detail");
                Assert.DoesNotContain("Text-box TOC appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertGeneratedEntries(toc, "Text-box TOC overview", "Text-box TOC detail");
                Assert.DoesNotContain("Text-box TOC appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesAppliesAnchorPageBreakBeforeToTextBoxHeadings() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesTextBoxHeadingAnchorBreak.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 2);
                document.AddParagraph("Intro text");
                Paragraph hostParagraph = CreateTextBoxHeadingParagraph("Anchored text-box heading", "Heading1");
                hostParagraph.ParagraphProperties ??= new ParagraphProperties();
                hostParagraph.ParagraphProperties.PageBreakBefore = new PageBreakBefore();
                AppendBodyParagraph(document, hostParagraph);

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                WordTableOfContentEntry entry = Assert.Single(report.Entries);
                Assert.Equal("Anchored text-box heading", entry.Text);
                Assert.Equal(2, entry.PageNumber);
                Assert.Contains("Anchored text-box heading", TocText(toc));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesSeparatesParentAndTextBoxHeadingText() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesParentAndTextBoxHeadings.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 2);
                AppendBodyParagraph(document, CreateParentAndTextBoxHeadingParagraph("Parent TOC heading", "Anchored child heading"));

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(new[] { "Anchored child heading", "Parent TOC heading" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.DoesNotContain(report.Entries, entry => entry.Text.Contains("Parent TOC headingAnchored child heading", StringComparison.Ordinal));
                Assert.Contains("Parent TOC heading", TocText(toc));
                Assert.Contains("Anchored child heading", TocText(toc));
                Assert.DoesNotContain("Parent TOC headingAnchored child heading", TocText(toc));
                WordTableOfContentEntry parentEntry = report.Entries.Single(entry => entry.Text == "Parent TOC heading");
                BookmarkStart parentBookmark = Assert.Single(document._document.Body!.Descendants<BookmarkStart>(), bookmark => bookmark.Name == parentEntry.BookmarkName);
                Assert.False(parentBookmark.Ancestors<TextBoxContent>().Any());
                Assert.Equal(2, report.Entries.Select(entry => entry.BookmarkName).Distinct(StringComparer.Ordinal).Count());
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesSupportsWordGeneratedTextBoxHeadings() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-toc-text-box.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated text-box TOC fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesWordGeneratedTextBox.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Equal(1, toc.MinLevel);
                Assert.Equal(2, toc.MaxLevel);
                Assert.Contains(document._document.Body!.Descendants<TextBoxContent>(), textBox =>
                    textBox.Descendants<Paragraph>().Any(paragraph =>
                        paragraph.ParagraphProperties?.ParagraphStyleId?.Val == "Heading2" &&
                        paragraph.Descendants<Text>().Any(text => text.Text.Contains("Word-generated text-box TOC detail", StringComparison.Ordinal))));
                Assert.Contains("Word-generated text-box TOC overview", TocText(toc));
                Assert.Contains("Word-generated text-box TOC detail", TocText(toc));
                Assert.DoesNotContain("Word-generated text-box TOC appendix", TocText(toc));

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(1, report.SkippedHeadingCount);
                Assert.Equal(new[] { "Word-generated text-box TOC overview", "Word-generated text-box TOC detail" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                WordTableOfContentEntry detailEntry = report.Entries.Single(entry => entry.Text == "Word-generated text-box TOC detail");
                Assert.StartsWith("_Toc", detailEntry.BookmarkName);
                Assert.Contains(document._document.Body!.Descendants<TextBoxContent>(), textBox =>
                    textBox.Descendants<BookmarkStart>().Any(bookmark => bookmark.Name == detailEntry.BookmarkName));
                AssertGeneratedEntries(toc, "Word-generated text-box TOC overview", "Word-generated text-box TOC detail");
                Assert.DoesNotContain("Word-generated text-box TOC appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(toc.SdtBlock.Descendants<SimpleField>().Single().Dirty?.Value);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertGeneratedEntries(toc, "Word-generated text-box TOC overview", "Word-generated text-box TOC detail");
                Assert.DoesNotContain("Word-generated text-box TOC appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesSupportsWordGeneratedTableTextBoxHeadings() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-toc-table-text-box.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated table text-box TOC fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesWordGeneratedTableTextBox.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Equal(1, toc.MinLevel);
                Assert.Equal(2, toc.MaxLevel);
                Assert.Contains(document._document.Body!.Descendants<Table>(), table =>
                    table.Descendants<TextBoxContent>().Any(textBox =>
                        textBox.Descendants<Paragraph>().Any(paragraph =>
                            paragraph.ParagraphProperties?.ParagraphStyleId?.Val == "Heading2" &&
                            paragraph.Descendants<Text>().Any(text => text.Text.Contains("Word-generated table text-box TOC detail", StringComparison.Ordinal)))));
                Assert.Contains("Word-generated table text-box TOC overview", TocText(toc));
                Assert.Contains("Word-generated table text-box TOC detail", TocText(toc));
                Assert.DoesNotContain("Word-generated table text-box TOC appendix", TocText(toc));

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(1, report.SkippedHeadingCount);
                Assert.Equal(new[] { "Word-generated table text-box TOC overview", "Word-generated table text-box TOC detail" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                WordTableOfContentEntry detailEntry = report.Entries.Single(entry => entry.Text == "Word-generated table text-box TOC detail");
                Assert.Contains(document._document.Body!.Descendants<Table>(), table =>
                    table.Descendants<TextBoxContent>().Any(textBox =>
                        textBox.Descendants<BookmarkStart>().Any(bookmark => bookmark.Name == detailEntry.BookmarkName)));
                AssertGeneratedEntries(toc, "Word-generated table text-box TOC overview", "Word-generated table text-box TOC detail");
                Assert.DoesNotContain("Word-generated table text-box TOC appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(toc.SdtBlock.Descendants<SimpleField>().Single().Dirty?.Value);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertGeneratedEntries(toc, "Word-generated table text-box TOC overview", "Word-generated table text-box TOC detail");
                Assert.DoesNotContain("Word-generated table text-box TOC appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesSupportsWordGeneratedOutlineLevels() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-outline-toc.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated outline-level TOC fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesWordGeneratedOutline.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Equal(1, toc.MinLevel);
                Assert.Equal(2, toc.MaxLevel);
                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("Word-generated outline root", TocText(toc));
                Assert.Contains("Word-generated outline child", TocText(toc));
                Assert.DoesNotContain("Word-generated outline appendix", TocText(toc));

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(1, report.SkippedHeadingCount);
                Assert.Equal(new[] { "Word-generated outline root", "Word-generated outline child" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                AssertGeneratedEntries(toc, "Word-generated outline root", "Word-generated outline child");
                Assert.DoesNotContain("Word-generated outline appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(toc.SdtBlock.Descendants<SimpleField>().Single().Dirty?.Value);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertGeneratedEntries(toc, "Word-generated outline root", "Word-generated outline child");
                Assert.DoesNotContain("Word-generated outline appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesSupportsWordGeneratedTcFields() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-tc-toc.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated TC TOC fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesWordGeneratedTc.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Equal(1, toc.MinLevel);
                Assert.Equal(3, toc.MaxLevel);
                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("Word-generated TC root", TocText(toc));
                Assert.Contains("Word-generated TC child", TocText(toc));
                Assert.DoesNotContain("Word-generated TC other type", TocText(toc));

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(1, report.SkippedHeadingCount);
                Assert.Equal(new[] { "Word-generated TC root", "Word-generated TC child" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                AssertGeneratedEntries(toc, "Word-generated TC root", "Word-generated TC child");
                Assert.DoesNotContain("Word-generated TC other type", TocText(toc));
                Assert.DoesNotContain("Word-generated TC appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(toc.SdtBlock.Descendants<SimpleField>().Single().Dirty?.Value);
                Assert.Contains(toc.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\f \"A\"", StringComparison.Ordinal));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertGeneratedEntries(toc, "Word-generated TC root", "Word-generated TC child");
                Assert.DoesNotContain("Word-generated TC other type", TocText(toc));
                Assert.DoesNotContain("Word-generated TC appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesHonorsTcOnlySourceSwitches() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesTcOnlySources.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 3);
                SetTocInstruction(toc, " TOC \\f \"A\" \\h \\z ");

                document.AddParagraph("Heading should stay out").SetStyle(WordParagraphStyles.Heading1);
                AppendBodyParagraph(document, new Paragraph(
                    new SimpleField { Instruction = " TC \"TC only entry\" \\f \"A\" \\l \"1\" " }));

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Single(report.Entries);
                Assert.Equal("TC only entry", report.Entries[0].Text);
                AssertGeneratedEntries(toc, "TC only entry");
                Assert.DoesNotContain("Heading should stay out", TocText(toc));
                Assert.Contains(toc.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\f \"A\"", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesHonorsCustomStyleOnlySourceSwitches() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesCustomStyleOnlySources.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 3);
                SetTocInstruction(toc, " TOC \\t \"CustomOnly,1\" \\h \\z ");

                document.AddParagraph("Heading should stay out").SetStyle(WordParagraphStyles.Heading1);
                AppendBodyParagraph(document, new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "CustomOnly" }),
                    new Run(new Text("Custom style entry") { Space = SpaceProcessingModeValues.Preserve })));

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Single(report.Entries);
                Assert.Equal("Custom style entry", report.Entries[0].Text);
                AssertGeneratedEntries(toc, "Custom style entry");
                Assert.DoesNotContain("Heading should stay out", TocText(toc));
                Assert.Contains(toc.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\t \"CustomOnly,1\"", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesIgnoresPageRefSimpleFieldsWhenFindingInstruction() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesIgnorePageRefSimpleFields.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 1);
                SdtContentBlock content = toc.SdtBlock.SdtContentBlock!;
                content.RemoveAllChildren();
                content.Append(new Paragraph(
                    new SimpleField(new Run(new Text("1"))) {
                        Instruction = " PAGEREF _TocGenerated \\h "
                    }));
                content.Append(CreateComplexFieldParagraph(" TOC \\o \"1-1\" \\h \\z ", "No table of contents entries found."));
                document.AddParagraph("Real TOC Heading").SetStyle(WordParagraphStyles.Heading1);

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Single(report.Entries);
                Assert.Equal("Real TOC Heading", report.Entries[0].Text);
                AssertGeneratedEntries(toc, "Real TOC Heading");
                SimpleField generatedField = toc.SdtBlock.Descendants<SimpleField>().Single(field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("TOC", StringComparison.OrdinalIgnoreCase));
                Assert.Contains("\\o \"1-1\"", generatedField.Instruction?.Value ?? generatedField.Instruction ?? string.Empty);
                Assert.DoesNotContain("PAGEREF", generatedField.Instruction?.Value ?? generatedField.Instruction ?? string.Empty, StringComparison.OrdinalIgnoreCase);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesSupportsWordGeneratedCustomStyleToc() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-custom-style-toc.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated custom-style TOC fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesWordGeneratedCustomStyle.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Equal(1, toc.MinLevel);
                Assert.Equal(3, toc.MaxLevel);
                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("No table of contents entries found.", TocText(toc));

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(1, report.SkippedHeadingCount);
                Assert.Equal(new[] { "Word-generated custom style root", "Word-generated custom style child" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                AssertGeneratedEntries(toc, "Word-generated custom style root", "Word-generated custom style child");
                Assert.DoesNotContain("Word-generated custom style excluded body", TocText(toc));
                Assert.DoesNotContain("Word-generated custom style appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.Contains(toc.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\t \"Custom TOC Root,1,Custom TOC Detail,2,Custom TOC Appendix,4\"", StringComparison.Ordinal));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertGeneratedEntries(toc, "Word-generated custom style root", "Word-generated custom style child");
                Assert.DoesNotContain("Word-generated custom style excluded body", TocText(toc));
                Assert.DoesNotContain("Word-generated custom style appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesHonorsImportedBookmarkScopeFilters() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesBookmarkScope.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 2);
                SetTocInstruction(toc, " TOC \\o \"1-3\" \\h \\z \\b \"ScopedToc\" ");

                document.AddParagraph("Outside before scope").SetStyle(WordParagraphStyles.Heading1);

                string bookmarkId = document.BookmarkId.ToString();
                AppendBodyParagraph(document, new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                    new BookmarkStart { Name = "ScopedToc", Id = bookmarkId },
                    new Run(new Text("Scoped root") { Space = SpaceProcessingModeValues.Preserve })));
                document.AddPageBreak();
                AppendBodyParagraph(document, new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "Heading2" }),
                    new Run(new Text("Scoped child") { Space = SpaceProcessingModeValues.Preserve }),
                    new BookmarkEnd { Id = bookmarkId }));

                document.AddParagraph("Outside after scope").SetStyle(WordParagraphStyles.Heading1);
                document.AddParagraph("Scoped appendix outside range").SetStyle(WordParagraphStyles.Heading3);

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedHeadingCount);
                Assert.Equal(new[] { "Scoped root", "Scoped child" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());

                string tocText = TocText(toc);
                Assert.Contains("Scoped root", tocText);
                Assert.Contains("Scoped child", tocText);
                Assert.DoesNotContain("Outside before scope", tocText);
                Assert.DoesNotContain("Outside after scope", tocText);
                Assert.DoesNotContain("Scoped appendix outside range", tocText);
                Assert.Contains(toc.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\b \"ScopedToc\"", StringComparison.Ordinal));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string tocText = TocText(toc);

                Assert.Contains("Scoped root", tocText);
                Assert.Contains("Scoped child", tocText);
                Assert.DoesNotContain("Outside before scope", tocText);
                Assert.DoesNotContain("Outside after scope", tocText);
                Assert.Contains(toc.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\b \"ScopedToc\"", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesSupportsWordGeneratedBookmarkScopeToc() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-bookmark-scope-toc.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated bookmark-scope TOC fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesWordGeneratedBookmarkScope.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Equal(1, toc.MinLevel);
                Assert.Equal(3, toc.MaxLevel);
                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("Word-generated scoped TOC root", TocText(toc));
                Assert.Contains("Word-generated scoped TOC child", TocText(toc));
                Assert.DoesNotContain("Word-generated scoped TOC outside before", TocText(toc));
                Assert.DoesNotContain("Word-generated scoped TOC outside after", TocText(toc));

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedHeadingCount);
                Assert.Equal(new[] { "Word-generated scoped TOC root", "Word-generated scoped TOC child" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                AssertGeneratedEntries(toc, "Word-generated scoped TOC root", "Word-generated scoped TOC child");
                Assert.DoesNotContain("Word-generated scoped TOC outside before", TocText(toc));
                Assert.DoesNotContain("Word-generated scoped TOC outside after", TocText(toc));
                Assert.DoesNotContain("Word-generated scoped TOC appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.Contains(toc.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\b \"OfficeIMO_TocScope\"", StringComparison.Ordinal));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertGeneratedEntries(toc, "Word-generated scoped TOC root", "Word-generated scoped TOC child");
                Assert.DoesNotContain("Word-generated scoped TOC outside before", TocText(toc));
                Assert.DoesNotContain("Word-generated scoped TOC outside after", TocText(toc));
                Assert.DoesNotContain("Word-generated scoped TOC appendix", TocText(toc));
                Assert.Equal(2, toc.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesHonorsImportedPageNumberSuppression() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesPageNumberSuppression.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 3);
                SetTocInstruction(toc, " TOC \\o \"1-3\" \\h \\z \\n \"2-3\" ");

                document.AddParagraph("Visible page root").SetStyle(WordParagraphStyles.Heading1);
                document.AddPageBreak();
                document.AddParagraph("Suppressed page child").SetStyle(WordParagraphStyles.Heading2);
                document.AddParagraph("Suppressed page detail").SetStyle(WordParagraphStyles.Heading3);

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(3, report.EntryCount);
                Assert.Equal(new[] { 1, 2, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                AssertTocEntryPageNumberState(toc, "Visible page root", expectedStyleId: "TOC1", shouldContainPageNumber: true);
                AssertTocEntryPageNumberState(toc, "Suppressed page child", expectedStyleId: "TOC2", shouldContainPageNumber: false);
                AssertTocEntryPageNumberState(toc, "Suppressed page detail", expectedStyleId: "TOC3", shouldContainPageNumber: false);
                Assert.Contains(toc.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\n \"2-3\"", StringComparison.Ordinal));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertTocEntryPageNumberState(toc, "Visible page root", expectedStyleId: "TOC1", shouldContainPageNumber: true);
                AssertTocEntryPageNumberState(toc, "Suppressed page child", expectedStyleId: "TOC2", shouldContainPageNumber: false);
                AssertTocEntryPageNumberState(toc, "Suppressed page detail", expectedStyleId: "TOC3", shouldContainPageNumber: false);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesHonorsImportedPageNumberSuppressionForAllLevels() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesAllPageNumbersSuppressed.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 2);
                SetTocInstruction(toc, " TOC \\o \"1-2\" \\h \\z \\n ");

                document.AddParagraph("All suppressed root").SetStyle(WordParagraphStyles.Heading1);
                document.AddPageBreak();
                document.AddParagraph("All suppressed child").SetStyle(WordParagraphStyles.Heading2);

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                AssertTocEntryPageNumberState(toc, "All suppressed root", expectedStyleId: "TOC1", shouldContainPageNumber: false);
                AssertTocEntryPageNumberState(toc, "All suppressed child", expectedStyleId: "TOC2", shouldContainPageNumber: false);
                Assert.Contains(toc.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\n", StringComparison.Ordinal));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertTocEntryPageNumberState(toc, "All suppressed root", expectedStyleId: "TOC1", shouldContainPageNumber: false);
                AssertTocEntryPageNumberState(toc, "All suppressed child", expectedStyleId: "TOC2", shouldContainPageNumber: false);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesSupportsWordGeneratedPageNumberSuppressionToc() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-page-number-suppression-toc.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated page-number suppression TOC fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesWordGeneratedPageNumberSuppression.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Equal(1, toc.MinLevel);
                Assert.Equal(3, toc.MaxLevel);
                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("Word-generated no-page root", TocText(toc));
                Assert.Contains("Word-generated no-page child", TocText(toc));
                Assert.Contains("Word-generated no-page detail", TocText(toc));

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(3, report.EntryCount);
                Assert.Equal(new[] { "Word-generated no-page root", "Word-generated no-page child", "Word-generated no-page detail" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                AssertTocEntryPageNumberState(toc, "Word-generated no-page root", expectedStyleId: "TOC1", shouldContainPageNumber: true);
                AssertTocEntryPageNumberState(toc, "Word-generated no-page child", expectedStyleId: "TOC2", shouldContainPageNumber: false);
                AssertTocEntryPageNumberState(toc, "Word-generated no-page detail", expectedStyleId: "TOC3", shouldContainPageNumber: false);
                Assert.Contains(toc.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\n \"2-3\"", StringComparison.Ordinal));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertTocEntryPageNumberState(toc, "Word-generated no-page root", expectedStyleId: "TOC1", shouldContainPageNumber: true);
                AssertTocEntryPageNumberState(toc, "Word-generated no-page child", expectedStyleId: "TOC2", shouldContainPageNumber: false);
                AssertTocEntryPageNumberState(toc, "Word-generated no-page detail", expectedStyleId: "TOC3", shouldContainPageNumber: false);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesSupportsWordGeneratedPageNumberSeparatorToc() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-page-number-separator-toc.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated page-number separator TOC fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesWordGeneratedPageNumberSeparator.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Equal(1, toc.MinLevel);
                Assert.Equal(3, toc.MaxLevel);
                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("Word-generated no-page root", TocText(toc));
                Assert.Contains("Word-generated no-page child", TocText(toc));
                Assert.Contains("Word-generated no-page detail", TocText(toc));

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(3, report.EntryCount);
                Assert.Equal(new[] { "Word-generated no-page root", "Word-generated no-page child", "Word-generated no-page detail" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                AssertTocEntryPageNumberSeparator(toc, "Word-generated no-page root", expectedStyleId: "TOC1", expectedSeparator: " -> ", expectedPageNumber: "1");
                AssertTocEntryPageNumberSeparator(toc, "Word-generated no-page child", expectedStyleId: "TOC2", expectedSeparator: " -> ", expectedPageNumber: "2");
                AssertTocEntryPageNumberSeparator(toc, "Word-generated no-page detail", expectedStyleId: "TOC3", expectedSeparator: " -> ", expectedPageNumber: "2");
                Assert.Contains(toc.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\p \" -> \"", StringComparison.Ordinal));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertTocEntryPageNumberSeparator(toc, "Word-generated no-page root", expectedStyleId: "TOC1", expectedSeparator: " -> ", expectedPageNumber: "1");
                AssertTocEntryPageNumberSeparator(toc, "Word-generated no-page child", expectedStyleId: "TOC2", expectedSeparator: " -> ", expectedPageNumber: "2");
                AssertTocEntryPageNumberSeparator(toc, "Word-generated no-page detail", expectedStyleId: "TOC3", expectedSeparator: " -> ", expectedPageNumber: "2");
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshEntriesHonorsImportedPageNumberSeparator() {
            string filePath = Path.Combine(_directoryWithFiles, "TocRefreshEntriesPageNumberSeparator.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 2);
                SetTocInstruction(toc, " TOC \\o \"1-2\" \\h \\z \\p \" -> \" ");

                document.AddParagraph("Separator root").SetStyle(WordParagraphStyles.Heading1);
                document.AddPageBreak();
                document.AddParagraph("Separator child").SetStyle(WordParagraphStyles.Heading2);

                WordTableOfContentRefreshReport report = toc.RefreshEntries();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                AssertTocEntryPageNumberSeparator(toc, "Separator root", expectedStyleId: "TOC1", expectedSeparator: " -> ", expectedPageNumber: "1");
                AssertTocEntryPageNumberSeparator(toc, "Separator child", expectedStyleId: "TOC2", expectedSeparator: " -> ", expectedPageNumber: "2");
                Assert.Contains(toc.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\p \" -> \"", StringComparison.Ordinal));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent toc = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertTocEntryPageNumberSeparator(toc, "Separator root", expectedStyleId: "TOC1", expectedSeparator: " -> ", expectedPageNumber: "1");
                AssertTocEntryPageNumberSeparator(toc, "Separator child", expectedStyleId: "TOC2", expectedSeparator: " -> ", expectedPageNumber: "2");
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresGeneratesCaptionEntriesAndBookmarks() {
            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFigures.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent list = document.AddTableOfContent();

                AddGeneratedCaptionParagraph(document, "FigureNetwork", "Figure", "Network diagram");
                document.AddPageBreak();
                AddGeneratedCaptionParagraph(document, "FigureLatency", "Figure", "Latency chart");
                AddGeneratedCaptionParagraph(document, "TableSignals", "Table", "Signal summary");

                document.UpdateFieldsAndGetReport();
                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(2, report.EntryCount);
                Assert.Equal(new[] { "Figure 1 Network diagram", "Figure 2 Latency chart" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.All(report.Entries, entry => Assert.StartsWith("Figure", entry.Text));
                Assert.All(report.Entries, entry => Assert.Contains(document._document.Body!.Descendants<BookmarkStart>(), bookmark => bookmark.Name == entry.BookmarkName));
                Assert.Contains("List of Figures", TocText(list));
                Assert.DoesNotContain("Signal summary", TocText(list));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Network diagram", TocText(list));
                Assert.Contains("Figure 2 Latency chart", TocText(list));
                Assert.DoesNotContain("Signal summary", TocText(list));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshCaptionListCountsEveryExplicitPageBreak() {
            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresMultipleBreaks.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent list = document.AddTableOfContent();
                AddGeneratedCaptionParagraph(document, "FigureOne", "Figure", "First diagram");
                AppendBodyParagraph(document, new Paragraph(
                    new Run(new Break { Type = BreakValues.Page }),
                    new Run(new Break { Type = BreakValues.Page })));
                AddGeneratedCaptionParagraph(document, "FigureTwo", "Figure", "Second diagram");

                document.UpdateFieldsAndGetReport();
                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal(new[] { "Figure 1 First diagram", "Figure 2 Second diagram" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 3 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshCaptionListHonorsImportedPageNumberSeparator() {
            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresPageNumberSeparator.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent list = document.AddTableOfContent();
                SetTocInstruction(list, " TOC \\h \\z \\c \"Figure\" \\p \" :: \" ");

                AddGeneratedCaptionParagraph(document, "FigureSeparatorNetwork", "Figure", "Separator network diagram");
                document.AddPageBreak();
                AddGeneratedCaptionParagraph(document, "FigureSeparatorLatency", "Figure", "Separator latency chart");
                AddGeneratedCaptionParagraph(document, "TableSeparatorExcluded", "Table", "Separator excluded table");

                document.UpdateFieldsAndGetReport();
                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(2, report.EntryCount);
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                AssertTocEntryPageNumberSeparator(list, "Figure 1 Separator network diagram", expectedStyleId: "TOC1", expectedSeparator: " :: ", expectedPageNumber: "1");
                AssertTocEntryPageNumberSeparator(list, "Figure 2 Separator latency chart", expectedStyleId: "TOC1", expectedSeparator: " :: ", expectedPageNumber: "2");
                Assert.DoesNotContain("Separator excluded table", TocText(list));
                Assert.Contains(list.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\p \" :: \"", StringComparison.Ordinal));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertTocEntryPageNumberSeparator(list, "Figure 1 Separator network diagram", expectedStyleId: "TOC1", expectedSeparator: " :: ", expectedPageNumber: "1");
                AssertTocEntryPageNumberSeparator(list, "Figure 2 Separator latency chart", expectedStyleId: "TOC1", expectedSeparator: " :: ", expectedPageNumber: "2");
                Assert.DoesNotContain("Separator excluded table", TocText(list));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshCaptionListHonorsImportedPageNumberSuppression() {
            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresPageNumbersSuppressed.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent list = document.AddTableOfContent();
                SetTocInstruction(list, " TOC \\h \\z \\c \"Figure\" \\n ");

                AddGeneratedCaptionParagraph(document, "FigureNoPageNetwork", "Figure", "No-page network diagram");
                document.AddPageBreak();
                AddGeneratedCaptionParagraph(document, "FigureNoPageLatency", "Figure", "No-page latency chart");
                AddGeneratedCaptionParagraph(document, "TableStillExcluded", "Table", "Still excluded");

                document.UpdateFieldsAndGetReport();
                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(2, report.EntryCount);
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains(list.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\n", StringComparison.Ordinal));
                AssertTocEntryPageNumberState(list, "Figure 1 No-page network diagram", expectedStyleId: "TOC1", shouldContainPageNumber: false);
                AssertTocEntryPageNumberState(list, "Figure 2 No-page latency chart", expectedStyleId: "TOC1", shouldContainPageNumber: false);
                Assert.DoesNotContain("Still excluded", TocText(list));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertTocEntryPageNumberState(list, "Figure 1 No-page network diagram", expectedStyleId: "TOC1", shouldContainPageNumber: false);
                AssertTocEntryPageNumberState(list, "Figure 2 No-page latency chart", expectedStyleId: "TOC1", shouldContainPageNumber: false);
                Assert.Contains(list.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\n", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresSupportsWordGeneratedCaptions() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-list-of-figures.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated list-of-figures fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresWordGenerated.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("Word-authored network diagram", TocText(list));
                Assert.Contains("Word-authored latency chart", TocText(list));

                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedCaptionCount);
                Assert.Equal(new[] { "Figure 1 Word-authored network diagram", "Figure 2 Word-authored latency chart" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Word-authored network diagram", TocText(list));
                Assert.Contains("Figure 2 Word-authored latency chart", TocText(list));
                Assert.DoesNotContain("Word-authored signal table", TocText(list));
                Assert.All(report.Entries, entry => Assert.Contains(document._document.Body!.Descendants<BookmarkStart>(), bookmark => bookmark.Name == entry.BookmarkName));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Word-authored network diagram", TocText(list));
                Assert.Contains("Figure 2 Word-authored latency chart", TocText(list));
                Assert.DoesNotContain("Word-authored signal table", TocText(list));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresSupportsHeaderFooterCaptions() {
            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresHeaderFooter.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent list = document.AddTableOfContent();
                AddGeneratedCaptionParagraph(document, "_BodyCaption", "Figure", "1", "Body deployment view");
                AppendBodyParagraph(document, new Paragraph(
                    new Run(new Break { Type = BreakValues.Page }),
                    new Run(new Break { Type = BreakValues.Page })));
                document.AddHeadersAndFooters();

                Header header = document._wordprocessingDocument.MainDocumentPart!.HeaderParts.Single().Header!;
                Footer footer = document._wordprocessingDocument.MainDocumentPart!.FooterParts.Single().Footer!;
                AppendCaptionParagraph(header, "Figure", "2", "Header architecture map");
                AppendCaptionParagraph(footer, "Figure", "3", "Footer recovery map");

                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(3, report.EntryCount);
                Assert.Equal(0, report.SkippedCaptionCount);
                Assert.Equal(
                    new[] {
                        "Figure 1 Body deployment view",
                        "Figure 2 Header architecture map",
                        "Figure 3 Footer recovery map"
                    },
                    report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 1, 1 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Body deployment view", TocText(list));
                Assert.Contains("Figure 2 Header architecture map", TocText(list));
                Assert.Contains("Figure 3 Footer recovery map", TocText(list));
                Assert.Contains(header.Descendants<BookmarkStart>(), bookmark => report.Entries.Any(entry => entry.BookmarkName == bookmark.Name));
                Assert.Contains(footer.Descendants<BookmarkStart>(), bookmark => report.Entries.Any(entry => entry.BookmarkName == bookmark.Name));
                Assert.Equal(3, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Body deployment view", TocText(list));
                Assert.Contains("Figure 2 Header architecture map", TocText(list));
                Assert.Contains("Figure 3 Footer recovery map", TocText(list));
                Assert.Equal(3, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresSupportsFootnoteEndnoteCaptions() {
            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresFootnoteEndnote.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent list = document.AddTableOfContent();
                AddGeneratedCaptionParagraph(document, "_BodyCaption", "Figure", "1", "Body deployment view");
                document.AddParagraph("Footnote caption anchor").AddFootNote("Footnote body");
                document.AddParagraph("Endnote caption anchor").AddEndNote("Endnote body");

                Footnote footnote = document._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!.Elements<Footnote>().First(note => note.Type == null);
                Endnote endnote = document._wordprocessingDocument.MainDocumentPart!.EndnotesPart!.Endnotes!.Elements<Endnote>().First(note => note.Type == null);
                AppendCaptionParagraph(footnote, "Figure", "2", "Footnote architecture map");
                AppendCaptionParagraph(endnote, "Figure", "3", "Endnote recovery map");

                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(3, report.EntryCount);
                Assert.Equal(0, report.SkippedCaptionCount);
                Assert.Equal(
                    new[] {
                        "Figure 1 Body deployment view",
                        "Figure 2 Footnote architecture map",
                        "Figure 3 Endnote recovery map"
                    },
                    report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 1, 1 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Body deployment view", TocText(list));
                Assert.Contains("Figure 2 Footnote architecture map", TocText(list));
                Assert.Contains("Figure 3 Endnote recovery map", TocText(list));
                Assert.Contains(footnote.Descendants<BookmarkStart>(), bookmark => report.Entries.Any(entry => entry.BookmarkName == bookmark.Name));
                Assert.Contains(endnote.Descendants<BookmarkStart>(), bookmark => report.Entries.Any(entry => entry.BookmarkName == bookmark.Name));
                Assert.Equal(3, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Body deployment view", TocText(list));
                Assert.Contains("Figure 2 Footnote architecture map", TocText(list));
                Assert.Contains("Figure 3 Endnote recovery map", TocText(list));
                Assert.Equal(3, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresSupportsWordGeneratedTableCellCaptions() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-list-of-figures-table-cell.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated list-of-figures table-cell fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresWordGeneratedTableCell.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Contains(document._document.Body!.Descendants<Table>(), table =>
                    table.Descendants<Text>().Any(text => text.Text.Contains("Word-authored latency chart", StringComparison.Ordinal)));

                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedCaptionCount);
                Assert.Equal(new[] { "Figure 1 Word-authored network diagram", "Figure 2 Word-authored latency chart" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Word-authored network diagram", TocText(list));
                Assert.Contains("Figure 2 Word-authored latency chart", TocText(list));
                Assert.DoesNotContain("Word-authored signal table", TocText(list));
                Assert.All(report.Entries, entry => Assert.Contains(document._document.Body!.Descendants<BookmarkStart>(), bookmark => bookmark.Name == entry.BookmarkName));
                Assert.Contains(document._document.Body!.Descendants<Table>(), table =>
                    table.Descendants<BookmarkStart>().Any(bookmark => report.Entries.Any(entry => entry.BookmarkName == bookmark.Name)));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Word-authored network diagram", TocText(list));
                Assert.Contains("Figure 2 Word-authored latency chart", TocText(list));
                Assert.DoesNotContain("Word-authored signal table", TocText(list));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresSupportsWordGeneratedContentControlCaptions() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-list-of-figures-content-control.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated list-of-figures content-control fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresWordGeneratedContentControl.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Contains(document._document.Body!.Descendants<SdtBlock>(), control =>
                    control.Descendants<Text>().Any(text => text.Text.Contains("Word-authored latency chart", StringComparison.Ordinal)));

                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedCaptionCount);
                Assert.Equal(new[] { "Figure 1 Word-authored network diagram", "Figure 2 Word-authored latency chart" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Word-authored network diagram", TocText(list));
                Assert.Contains("Figure 2 Word-authored latency chart", TocText(list));
                Assert.DoesNotContain("Word-authored signal table", TocText(list));
                Assert.All(report.Entries, entry => Assert.Contains(document._document.Body!.Descendants<BookmarkStart>(), bookmark => bookmark.Name == entry.BookmarkName));
                Assert.Contains(document._document.Body!.Descendants<SdtBlock>(), control =>
                    control.Descendants<BookmarkStart>().Any(bookmark => report.Entries.Any(entry => entry.BookmarkName == bookmark.Name)));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Word-authored network diagram", TocText(list));
                Assert.Contains("Figure 2 Word-authored latency chart", TocText(list));
                Assert.DoesNotContain("Word-authored signal table", TocText(list));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresSupportsWordGeneratedTextBoxCaptions() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-list-of-figures-text-box.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated list-of-figures text-box fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresWordGeneratedTextBox.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Contains(document._document.Body!.Descendants<TextBoxContent>(), textBox =>
                    textBox.Descendants<Text>().Any(text => text.Text.Contains("Word-authored text-box diagram", StringComparison.Ordinal)));

                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedCaptionCount);
                Assert.Equal(new[] { "Figure 1 Word-authored body diagram", "Figure 2 Word-authored text-box diagram" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Word-authored body diagram", TocText(list));
                Assert.Contains("Figure 2 Word-authored text-box diagram", TocText(list));
                Assert.DoesNotContain("Word-authored excluded table", TocText(list));
                Assert.All(report.Entries, entry => Assert.Contains(document._document.Body!.Descendants<BookmarkStart>(), bookmark => bookmark.Name == entry.BookmarkName));
                Assert.Contains(document._document.Body!.Descendants<TextBoxContent>(), textBox =>
                    textBox.Descendants<BookmarkStart>().Any(bookmark => report.Entries.Any(entry => entry.BookmarkName == bookmark.Name)));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Word-authored body diagram", TocText(list));
                Assert.Contains("Figure 2 Word-authored text-box diagram", TocText(list));
                Assert.DoesNotContain("Word-authored excluded table", TocText(list));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresSupportsWordGeneratedTableTextBoxCaptions() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-list-of-figures-table-text-box.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated list-of-figures table text-box fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresWordGeneratedTableTextBox.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Contains(document._document.Body!.Descendants<Table>(), table =>
                    table.Descendants<TextBoxContent>().Any(textBox =>
                        textBox.Descendants<Text>().Any(text => text.Text.Contains("Word-authored table text-box diagram", StringComparison.Ordinal))));

                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedCaptionCount);
                Assert.Equal(new[] { "Figure 1 Word-authored body diagram", "Figure 2 Word-authored table text-box diagram" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Word-authored body diagram", TocText(list));
                Assert.Contains("Figure 2 Word-authored table text-box diagram", TocText(list));
                Assert.All(report.Entries, entry => Assert.Contains(document._document.Body!.Descendants<BookmarkStart>(), bookmark => bookmark.Name == entry.BookmarkName));
                Assert.Contains(document._document.Body!.Descendants<Table>(), table =>
                    table.Descendants<TextBoxContent>().Any(textBox =>
                        textBox.Descendants<BookmarkStart>().Any(bookmark => report.Entries.Any(entry => entry.BookmarkName == bookmark.Name))));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Word-authored body diagram", TocText(list));
                Assert.Contains("Figure 2 Word-authored table text-box diagram", TocText(list));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresSupportsTextBoxCaptions() {
            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresTextBox.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent list = document.AddTableOfContent();

                AddGeneratedCaptionParagraph(document, "FigureBodyTopology", "Figure", "Body topology map");
                document.AddPageBreak();
                AddTextBoxCaptionParagraph(document, "Figure", "Text-box architecture sketch");
                AddGeneratedCaptionParagraph(document, "TableTextBoxExcluded", "Table", "Excluded table");

                document.UpdateFieldsAndGetReport();
                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedCaptionCount);
                Assert.Equal(new[] { "Figure 1 Body topology map", "Figure 2 Text-box architecture sketch" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains(document._document.Body!.Descendants<V.TextBox>(), textBox =>
                    textBox.Descendants<Text>().Any(text => text.Text.Contains("Text-box architecture sketch", StringComparison.Ordinal)));
                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Body topology map", TocText(list));
                Assert.Contains("Figure 2 Text-box architecture sketch", TocText(list));
                Assert.DoesNotContain("Excluded table", TocText(list));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Body topology map", TocText(list));
                Assert.Contains("Figure 2 Text-box architecture sketch", TocText(list));
                Assert.DoesNotContain("Excluded table", TocText(list));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresSupportsWordGeneratedPageNumberSeparator() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-list-of-figures-separator.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated list-of-figures page-number separator fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresWordGeneratedPageNumberSeparator.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("Word-authored network diagram", TocText(list));
                Assert.Contains("Word-authored latency chart", TocText(list));

                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedCaptionCount);
                Assert.Equal(new[] { "Figure 1 Word-authored network diagram", "Figure 2 Word-authored latency chart" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains(list.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\p \" :: \"", StringComparison.Ordinal));
                AssertTocEntryPageNumberSeparator(list, "Figure 1 Word-authored network diagram", expectedStyleId: "TOC1", expectedSeparator: " :: ", expectedPageNumber: "1");
                AssertTocEntryPageNumberSeparator(list, "Figure 2 Word-authored latency chart", expectedStyleId: "TOC1", expectedSeparator: " :: ", expectedPageNumber: "2");
                Assert.DoesNotContain("Word-authored signal table", TocText(list));
                Assert.All(report.Entries, entry => Assert.Contains(document._document.Body!.Descendants<BookmarkStart>(), bookmark => bookmark.Name == entry.BookmarkName));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertTocEntryPageNumberSeparator(list, "Figure 1 Word-authored network diagram", expectedStyleId: "TOC1", expectedSeparator: " :: ", expectedPageNumber: "1");
                AssertTocEntryPageNumberSeparator(list, "Figure 2 Word-authored latency chart", expectedStyleId: "TOC1", expectedSeparator: " :: ", expectedPageNumber: "2");
                Assert.DoesNotContain("Word-authored signal table", TocText(list));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresSupportsWordGeneratedPageNumberSuppression() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-list-of-figures-no-page-numbers.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated list-of-figures page-number suppression fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresWordGeneratedNoPageNumbers.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("Word-authored no-page network diagram", TocText(list));
                Assert.Contains("Word-authored no-page latency chart", TocText(list));

                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedCaptionCount);
                Assert.Equal(new[] { "Figure 1 Word-authored no-page network diagram", "Figure 2 Word-authored no-page latency chart" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains(list.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\n", StringComparison.Ordinal));
                AssertTocEntryPageNumberState(list, "Figure 1 Word-authored no-page network diagram", expectedStyleId: "TOC1", shouldContainPageNumber: false);
                AssertTocEntryPageNumberState(list, "Figure 2 Word-authored no-page latency chart", expectedStyleId: "TOC1", shouldContainPageNumber: false);
                Assert.DoesNotContain("Word-authored excluded table", TocText(list));
                Assert.All(report.Entries, entry => Assert.Contains(document._document.Body!.Descendants<BookmarkStart>(), bookmark => bookmark.Name == entry.BookmarkName));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                AssertTocEntryPageNumberState(list, "Figure 1 Word-authored no-page network diagram", expectedStyleId: "TOC1", shouldContainPageNumber: false);
                AssertTocEntryPageNumberState(list, "Figure 2 Word-authored no-page latency chart", expectedStyleId: "TOC1", shouldContainPageNumber: false);
                Assert.DoesNotContain("Word-authored excluded table", TocText(list));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfTablesReplacesCaptionListContent() {
            string filePath = Path.Combine(_directoryWithFiles, "CaptionListTables.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent list = document.AddTableOfContent();

                AddGeneratedCaptionParagraph(document, "FigureNetwork", "Figure", "Network diagram");
                AddGeneratedCaptionParagraph(document, "TableSignals", "Table", "Signal summary");
                AddGeneratedCaptionParagraph(document, "TableAudit", "Table", "Audit detail");

                document.UpdateFieldsAndGetReport();
                Assert.Equal(1, list.RefreshListOfFigures().EntryCount);

                WordCaptionListRefreshReport report = list.RefreshListOfTables();

                Assert.Equal("Table", report.SequenceIdentifier);
                Assert.Equal(2, report.EntryCount);
                Assert.Equal(new[] { "Table 1 Signal summary", "Table 2 Audit detail" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Contains("List of Tables", TocText(list));
                Assert.DoesNotContain("Network diagram", TocText(list));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains("List of Tables", TocText(list));
                Assert.Contains("Table 1 Signal summary", TocText(list));
                Assert.Contains("Table 2 Audit detail", TocText(list));
                Assert.DoesNotContain("Network diagram", TocText(list));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfTablesSupportsWordGeneratedCaptions() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-list-of-tables.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated list-of-tables fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "CaptionListTablesWordGenerated.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("Word-authored signal table", TocText(list));
                Assert.Contains("Word-authored audit table", TocText(list));

                WordCaptionListRefreshReport report = list.RefreshListOfTables();

                Assert.Equal("Table", report.SequenceIdentifier);
                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedCaptionCount);
                Assert.Equal(new[] { "Table 1 Word-authored signal table", "Table 2 Word-authored audit table" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains("List of Tables", TocText(list));
                Assert.Contains("Table 1 Word-authored signal table", TocText(list));
                Assert.Contains("Table 2 Word-authored audit table", TocText(list));
                Assert.DoesNotContain("Word-authored network figure", TocText(list));
                Assert.All(report.Entries, entry => Assert.Contains(document._document.Body!.Descendants<BookmarkStart>(), bookmark => bookmark.Name == entry.BookmarkName));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains("List of Tables", TocText(list));
                Assert.Contains("Table 1 Word-authored signal table", TocText(list));
                Assert.Contains("Table 2 Word-authored audit table", TocText(list));
                Assert.DoesNotContain("Word-authored network figure", TocText(list));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshCaptionListSupportsWordGeneratedEquationCaptions() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-list-of-equations.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated list-of-equations fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "CaptionListEquationsWordGenerated.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("Word-authored quadratic identity", TocText(list));
                Assert.Contains("Word-authored integration identity", TocText(list));

                WordCaptionListRefreshReport report = list.RefreshCaptionList("Equation", "List of Equations");

                Assert.Equal("Equation", report.SequenceIdentifier);
                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedCaptionCount);
                Assert.Equal(new[] { "Equation 1 Word-authored quadratic identity", "Equation 2 Word-authored integration identity" }, report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains("List of Equations", TocText(list));
                Assert.Contains("Equation 1 Word-authored quadratic identity", TocText(list));
                Assert.Contains("Equation 2 Word-authored integration identity", TocText(list));
                Assert.DoesNotContain("Word-authored excluded figure", TocText(list));
                Assert.All(report.Entries, entry => Assert.Contains(document._document.Body!.Descendants<BookmarkStart>(), bookmark => bookmark.Name == entry.BookmarkName));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains("List of Equations", TocText(list));
                Assert.Contains("Equation 1 Word-authored quadratic identity", TocText(list));
                Assert.Contains("Equation 2 Word-authored integration identity", TocText(list));
                Assert.DoesNotContain("Word-authored excluded figure", TocText(list));
                Assert.Equal(2, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshCaptionListRejectsInvalidSequenceIdentifier() {
            string filePath = Path.Combine(_directoryWithFiles, "CaptionListInvalidSequence.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent list = document.AddTableOfContent();

                ArgumentException exception = Assert.Throws<ArgumentException>(() => list.RefreshCaptionList("Figure \\n"));

                Assert.Contains("Caption sequence identifier", exception.Message, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexGeneratesEntriesFromXeFields() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshEntries.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();

                AddIndexEntryParagraph(document, "Alpha topic", " XE \"Alpha\" ");
                document.AddPageBreak();
                AddIndexEntryParagraph(document, "Beta topic", " XE \"Beta\" ");
                document.AddPageBreak();
                AddIndexEntryParagraph(document, "Alpha follow-up", " XE \"Alpha\" ");
                AddIndexEntryParagraph(document, "Alpha details", " XE \"Alpha:Details\" ");
                AddIndexEntryParagraph(document, "Alpha deep detail", " XE \"Alpha:Details:Deep Dive\" ");
                AddIndexEntryParagraph(document, "Cross-reference index entry", " XE \"Gamma\" \\t \"See Beta\" ");
                AddIndexEntryParagraph(document, "Unsupported bookmark range", " XE \"Delta\" \\r \"BookmarkRange\" ");

                WordIndexRefreshReport report = index.RefreshIndex();

                Assert.Equal(5, report.EntryCount);
                Assert.Equal(1, report.SkippedEntryCount);
                Assert.Equal(new[] { "Alpha", "Alpha", "Alpha", "Beta", "Gamma" }, report.Entries.Select(entry => entry.Term).ToArray());
                Assert.Equal(new[] { null, "Details", "Details", null, null }, report.Entries.Select(entry => entry.Subterm).ToArray());
                Assert.Equal(new[] { "Details", "Deep Dive" }, report.Entries[2].Subterms);
                Assert.Equal(new[] { 1, 3 }, report.Entries[0].PageNumbers);
                Assert.Equal(new[] { 3 }, report.Entries[1].PageNumbers);
                Assert.Equal(new[] { 3 }, report.Entries[2].PageNumbers);
                Assert.Equal(new[] { 2 }, report.Entries[3].PageNumbers);
                Assert.True(report.Entries[4].IsCrossReference);
                Assert.Equal("See Beta", report.Entries[4].CrossReferenceText);
                Assert.Empty(report.Entries[4].PageNumbers);

                string indexText = TocText(index);
                Assert.Contains("Index", indexText);
                Assert.Contains("Alpha", indexText);
                Assert.Contains("1, 3", indexText);
                Assert.Contains("Details", indexText);
                Assert.Contains("Deep Dive", indexText);
                Assert.Contains("Beta", indexText);
                Assert.Contains("Gamma", indexText);
                Assert.Contains("See Beta", indexText);
                Assert.DoesNotContain("Delta", indexText);
                Assert.Contains(index.SdtBlock.Descendants<Paragraph>(), paragraph =>
                    paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Index3" &&
                    string.Concat(paragraph.Descendants<Text>().Select(text => text.Text)).Contains("Deep Dive", StringComparison.Ordinal));
                Assert.True(index.SdtBlock.Descendants<SimpleField>().Single(field => (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("INDEX", StringComparison.OrdinalIgnoreCase)).Dirty?.Value);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string indexText = TocText(index);

                Assert.Contains("Index", indexText);
                Assert.Contains("Alpha", indexText);
                Assert.Contains("1, 3", indexText);
                Assert.Contains("Details", indexText);
                Assert.Contains("Deep Dive", indexText);
                Assert.Contains("Beta", indexText);
                Assert.Contains("Gamma", indexText);
                Assert.Contains("See Beta", indexText);
                Assert.DoesNotContain("Delta", indexText);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresUsesNoteAnchorPageEstimates() {
            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresNoteAnchorPages.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent list = document.AddTableOfContent();
                AddGeneratedCaptionParagraph(document, "_BodyCaption", "Figure", "1", "Body deployment view");
                document.AddPageBreak();
                document.AddParagraph("Footnote caption anchor").AddFootNote("Footnote body");
                document.AddParagraph("Endnote caption anchor").AddEndNote("Endnote body");

                Footnote footnote = document._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!.Elements<Footnote>().First(note => note.Type == null);
                Endnote endnote = document._wordprocessingDocument.MainDocumentPart!.EndnotesPart!.Endnotes!.Elements<Endnote>().First(note => note.Type == null);
                AppendCaptionParagraph(footnote, "Figure", "2", "Footnote architecture map");
                AppendCaptionParagraph(endnote, "Figure", "3", "Endnote recovery map");

                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal(
                    new[] {
                        "Figure 1 Body deployment view",
                        "Figure 2 Footnote architecture map",
                        "Figure 3 Endnote recovery map"
                    },
                    report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 2, 2 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresIncludesParentAndAnchoredTextBoxCaptions() {
            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresParentAndTextBox.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent list = document.AddTableOfContent();
                AppendBodyParagraph(document, CreateParentAndTextBoxCaptionParagraph("Figure", "Parent deployment map", "Text-box fallback map"));

                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(2, report.EntryCount);
                Assert.Equal(
                    new[] { "Figure 1 Parent deployment map", "Figure 2 Text-box fallback map" },
                    report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 1 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains("Figure 1 Parent deployment map", TocText(list));
                Assert.Contains("Figure 2 Text-box fallback map", TocText(list));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresKeepsParentBookmarkOutsideTextBox() {
            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresParentTextBoxBookmark.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent list = document.AddTableOfContent();
                Paragraph host = CreateParentAndTextBoxCaptionParagraph("Figure", "Parent deployment map", "Text-box fallback map");
                Paragraph textBoxCaption = host.Descendants<TextBoxContent>().Single().Descendants<Paragraph>().Single();
                textBoxCaption.PrependChild(new BookmarkStart { Name = "_NestedTextBoxCaption", Id = "77" });
                textBoxCaption.Append(new BookmarkEnd { Id = "77" });
                AppendBodyParagraph(document, host);

                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal(2, report.EntryCount);
                Assert.Equal("Figure 1 Parent deployment map", report.Entries[0].Text);
                Assert.Equal("Figure 2 Text-box fallback map", report.Entries[1].Text);
                Assert.NotEqual("_NestedTextBoxCaption", report.Entries[0].BookmarkName);
                Assert.Equal("_NestedTextBoxCaption", report.Entries[1].BookmarkName);
                Assert.Contains(host.Elements<BookmarkStart>(), bookmark => bookmark.Name?.Value == report.Entries[0].BookmarkName);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexCountsEveryExplicitPageBreak() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshEntriesMultipleBreaks.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();
                AddIndexEntryParagraph(document, "Alpha topic", " XE \"Alpha\" ");
                AppendBodyParagraph(document, new Paragraph(
                    new Run(new Break { Type = BreakValues.Page }),
                    new Run(new Break { Type = BreakValues.Page })));
                AddIndexEntryParagraph(document, "Beta topic", " XE \"Beta\" ");

                WordIndexRefreshReport report = index.RefreshIndex("Generated Index");

                Assert.Equal(new[] { "Alpha", "Beta" }, report.Entries.Select(entry => entry.Term).ToArray());
                Assert.Equal(new[] { 1 }, report.Entries[0].PageNumbers);
                Assert.Equal(new[] { 3 }, report.Entries[1].PageNumbers);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexIgnoresExistingIndexBlockForPageEstimates() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshEntriesIgnoresExistingBlock.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();
                index.SdtBlock.SdtContentBlock!.Append(
                    new Paragraph(new Run(new Break { Type = BreakValues.Page })),
                    new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                AddIndexEntryParagraph(document, "Body topic", " XE \"BodyTopic\" ");

                WordIndexRefreshReport report = index.RefreshIndex("Generated Index");

                WordIndexEntry entry = Assert.Single(report.Entries);
                Assert.Equal("BodyTopic", entry.Term);
                Assert.Equal(new[] { 1 }, entry.PageNumbers);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexFindsXeFieldsInTablesAndContentControls() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshEntriesTablesContentControls.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();

                AddIndexEntryParagraph(document, "Body topic", " XE \"BodyTopic\" ");
                document.AddPageBreak();

                AppendBodyElement(document, new Table(
                    new TableProperties(new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto }),
                    new TableGrid(new GridColumn { Width = "5000" }),
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Width = "5000", Type = TableWidthUnitValues.Dxa }),
                            new Paragraph(
                                new Run(new Text("Table topic") { Space = SpaceProcessingModeValues.Preserve }),
                                CreateIndexEntryField(" XE \"TableTopic\" "))))));
                document.AddPageBreak();

                AppendBodyElement(document, new SdtBlock(
                    new SdtProperties(
                        new SdtAlias { Val = "Index Content Control Source" },
                        new Tag { Val = "index-content-control-source" }),
                    new SdtContentBlock(
                        new Paragraph(
                            new Run(new Text("Content control topic") { Space = SpaceProcessingModeValues.Preserve }),
                            CreateIndexEntryField(" XE \"ControlTopic\" ")))));

                WordIndexRefreshReport report = index.RefreshIndex("Container Index");

                Assert.Equal(3, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal(new[] { "BodyTopic", "ControlTopic", "TableTopic" }, report.Entries.Select(entry => entry.Term).ToArray());
                Assert.Equal(new[] { 1 }, report.Entries.Single(entry => entry.Term == "BodyTopic").PageNumbers);
                Assert.Equal(new[] { 2 }, report.Entries.Single(entry => entry.Term == "TableTopic").PageNumbers);
                Assert.Equal(new[] { 3 }, report.Entries.Single(entry => entry.Term == "ControlTopic").PageNumbers);

                string indexText = TocText(index);
                Assert.Contains("Container Index", indexText);
                Assert.Contains("BodyTopic", indexText);
                Assert.Contains("TableTopic", indexText);
                Assert.Contains("ControlTopic", indexText);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string indexText = TocText(index);

                Assert.Contains("Container Index", indexText);
                Assert.Contains("BodyTopic", indexText);
                Assert.Contains("TableTopic", indexText);
                Assert.Contains("ControlTopic", indexText);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexFindsXeFieldsInHeaderFooterAndNotes() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshEntriesHeaderFooterNotes.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();

                AddIndexEntryParagraph(document, "Body topic", " XE \"BodyTopic\" ");
                document.AddHeadersAndFooters();
                Header header = document._wordprocessingDocument.MainDocumentPart!.HeaderParts.Single().Header!;
                Footer footer = document._wordprocessingDocument.MainDocumentPart!.FooterParts.Single().Footer!;
                AppendIndexEntryParagraph(header, "Header topic", " XE \"HeaderTopic\" ");
                AppendIndexEntryParagraph(footer, "Footer topic", " XE \"FooterTopic\" ");

                document.AddParagraph("Footnote index anchor").AddFootNote("Footnote body");
                document.AddParagraph("Endnote index anchor").AddEndNote("Endnote body");
                Footnote footnote = document._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!.Elements<Footnote>().First(note => note.Type == null);
                Endnote endnote = document._wordprocessingDocument.MainDocumentPart!.EndnotesPart!.Endnotes!.Elements<Endnote>().First(note => note.Type == null);
                AppendIndexEntryParagraph(footnote, "Footnote topic", " XE \"FootnoteTopic\" ");
                AppendIndexEntryParagraph(endnote, "Endnote topic", " XE \"EndnoteTopic\" ");

                WordIndexRefreshReport report = index.RefreshIndex("Related-Part Index");

                Assert.Equal(5, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal(new[] { "BodyTopic", "EndnoteTopic", "FooterTopic", "FootnoteTopic", "HeaderTopic" }, report.Entries.Select(entry => entry.Term).ToArray());
                Assert.All(report.Entries, entry => Assert.Equal(new[] { 1 }, entry.PageNumbers));

                string indexText = TocText(index);
                Assert.Contains("Related-Part Index", indexText);
                Assert.Contains("BodyTopic", indexText);
                Assert.Contains("HeaderTopic", indexText);
                Assert.Contains("FooterTopic", indexText);
                Assert.Contains("FootnoteTopic", indexText);
                Assert.Contains("EndnoteTopic", indexText);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string indexText = TocText(index);

                Assert.Contains("Related-Part Index", indexText);
                Assert.Contains("HeaderTopic", indexText);
                Assert.Contains("FooterTopic", indexText);
                Assert.Contains("FootnoteTopic", indexText);
                Assert.Contains("EndnoteTopic", indexText);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexSupportsWordGeneratedComplexIndex() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-index.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated index fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshWordGenerated.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("WordAlpha", TocText(index));
                Assert.Contains("WordBeta", TocText(index));
                Assert.Contains("WordGamma", TocText(index));

                WordIndexRefreshReport report = index.RefreshIndex("Imported Index");

                Assert.Equal(4, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal(new[] { "WordAlpha", "WordAlpha", "WordBeta", "WordGamma" }, report.Entries.Select(entry => entry.Term).ToArray());
                Assert.Equal(new[] { null, "Detail", null, null }, report.Entries.Select(entry => entry.Subterm).ToArray());
                Assert.Equal(new[] { 1 }, report.Entries[0].PageNumbers);
                Assert.Equal(new[] { 2 }, report.Entries[1].PageNumbers);
                Assert.Equal(new[] { 2 }, report.Entries[2].PageNumbers);
                Assert.True(report.Entries[3].IsCrossReference);
                Assert.Equal("See WordBeta", report.Entries[3].CrossReferenceText);
                Assert.Empty(report.Entries[3].PageNumbers);

                string indexText = TocText(index);
                Assert.Contains("Imported Index", indexText);
                Assert.Contains("WordAlpha", indexText);
                Assert.Contains("Detail", indexText);
                Assert.Contains("WordBeta", indexText);
                Assert.Contains("WordGamma", indexText);
                Assert.Contains("See WordBeta", indexText);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string indexText = TocText(index);

                Assert.Contains("Imported Index", indexText);
                Assert.Contains("WordAlpha", indexText);
                Assert.Contains("Detail", indexText);
                Assert.Contains("WordBeta", indexText);
                Assert.Contains("WordGamma", indexText);
                Assert.Contains("See WordBeta", indexText);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexSupportsWordGeneratedContainerEntries() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-index-containers.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated index container fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshWordGeneratedContainers.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains(document._document.Body!.Descendants<Table>(), table =>
                    table.Descendants<FieldCode>().Any(code => (code.Text ?? string.Empty).Contains("WordTable", StringComparison.Ordinal)));
                Assert.Contains(document._document.Body!.Descendants<SdtBlock>(), control =>
                    control.Descendants<FieldCode>().Any(code => (code.Text ?? string.Empty).Contains("WordControl", StringComparison.Ordinal)));
                Assert.Contains("WordBody", TocText(index));
                Assert.Contains("WordTable", TocText(index));
                Assert.Contains("WordControl", TocText(index));

                WordIndexRefreshReport report = index.RefreshIndex("Imported Container Index");

                Assert.Equal(3, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal(new[] { "WordBody", "WordControl", "WordTable" }, report.Entries.Select(entry => entry.Term).ToArray());
                Assert.Equal(new[] { 1 }, report.Entries.Single(entry => entry.Term == "WordBody").PageNumbers);
                Assert.Equal(new[] { 2 }, report.Entries.Single(entry => entry.Term == "WordTable").PageNumbers);
                Assert.Equal(new[] { 3 }, report.Entries.Single(entry => entry.Term == "WordControl").PageNumbers);

                string indexText = TocText(index);
                Assert.Contains("Imported Container Index", indexText);
                Assert.Contains("WordBody", indexText);
                Assert.Contains("WordTable", indexText);
                Assert.Contains("WordControl", indexText);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string indexText = TocText(index);

                Assert.Contains("Imported Container Index", indexText);
                Assert.Contains("WordBody", indexText);
                Assert.Contains("WordTable", indexText);
                Assert.Contains("WordControl", indexText);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexSupportsWordGeneratedTextBoxEntries() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-index-text-box.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated index text-box fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshWordGeneratedTextBox.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains(document._document.Body!.Descendants<TextBoxContent>(), textBox =>
                    textBox.Descendants<FieldCode>().Any(code => (code.Text ?? string.Empty).Contains("WordTextBox", StringComparison.Ordinal)));

                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("WordBody", TocText(index));
                Assert.Contains("WordTextBox", TocText(index));

                WordIndexRefreshReport report = index.RefreshIndex("Imported Text Box Index");

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal(new[] { "WordBody", "WordTextBox" }, report.Entries.Select(entry => entry.Term).ToArray());
                Assert.Equal(new[] { 1 }, report.Entries.Single(entry => entry.Term == "WordBody").PageNumbers);
                Assert.Equal(new[] { 2 }, report.Entries.Single(entry => entry.Term == "WordTextBox").PageNumbers);

                string indexText = TocText(index);
                Assert.Contains("Imported Text Box Index", indexText);
                Assert.Contains("WordBody", indexText);
                Assert.Contains("WordTextBox", indexText);
                Assert.DoesNotContain("WordExcludedTail", indexText);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string indexText = TocText(index);

                Assert.Contains("Imported Text Box Index", indexText);
                Assert.Contains("WordBody", indexText);
                Assert.Contains("WordTextBox", indexText);
                Assert.DoesNotContain("WordExcludedTail", indexText);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexSupportsWordGeneratedTableTextBoxEntries() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-index-table-text-box.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated index table text-box fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshWordGeneratedTableTextBox.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains(document._document.Body!.Descendants<Table>(), table =>
                    table.Descendants<TextBoxContent>().Any(textBox =>
                        textBox.Descendants<FieldCode>().Any(code => (code.Text ?? string.Empty).Contains("WordTableTextBox", StringComparison.Ordinal))));

                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("WordBody", TocText(index));
                Assert.Contains("WordTableTextBox", TocText(index));

                WordIndexRefreshReport report = index.RefreshIndex("Imported Table Text Box Index");

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal(new[] { "WordBody", "WordTableTextBox" }, report.Entries.Select(entry => entry.Term).ToArray());
                Assert.Equal(new[] { 1 }, report.Entries.Single(entry => entry.Term == "WordBody").PageNumbers);
                Assert.Equal(new[] { 2 }, report.Entries.Single(entry => entry.Term == "WordTableTextBox").PageNumbers);

                string indexText = TocText(index);
                Assert.Contains("Imported Table Text Box Index", indexText);
                Assert.Contains("WordBody", indexText);
                Assert.Contains("WordTableTextBox", indexText);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string indexText = TocText(index);

                Assert.Contains("Imported Table Text Box Index", indexText);
                Assert.Contains("WordBody", indexText);
                Assert.Contains("WordTableTextBox", indexText);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexSupportsWordGeneratedBookmarkPageRange() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-index-page-range.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated index page-range fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshWordGeneratedPageRange.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("WordRange", TocText(index));

                WordIndexRefreshReport report = index.RefreshIndex("Imported Range Index");

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal(new[] { "WordLoose", "WordRange" }, report.Entries.Select(entry => entry.Term).ToArray());
                Assert.Equal(new[] { 2 }, report.Entries[0].PageNumbers);
                Assert.Equal(new[] { "2" }, report.Entries[0].PageReferences);
                Assert.Empty(report.Entries[1].PageNumbers);
                Assert.Equal(new[] { "1-2" }, report.Entries[1].PageReferences);
                Assert.Equal("1-2", report.Entries[1].PageNumbersText);

                string indexText = TocText(index);
                Assert.Contains("Imported Range Index", indexText);
                Assert.Contains("WordLoose", indexText);
                Assert.Contains("WordRange", indexText);
                Assert.Contains("1-2", indexText);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string indexText = TocText(index);

                Assert.Contains("Imported Range Index", indexText);
                Assert.Contains("WordLoose", indexText);
                Assert.Contains("WordRange", indexText);
                Assert.Contains("1-2", indexText);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexSupportsWordGeneratedCustomSeparators() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-index-custom-separators.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-generated index custom-separator fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshWordGeneratedCustomSeparators.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.DoesNotContain(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));

                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains(document._document.Body!.Elements<SdtBlock>(), block =>
                    block.Descendants<DocPartGallery>().Any(gallery => gallery.Val == "Table of Contents"));
                Assert.Contains("WordLoose", TocText(index));
                Assert.Contains("WordRange", TocText(index));

                WordIndexRefreshReport report = index.RefreshIndex("Imported Separator Index");

                Assert.Equal(3, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal("2 | 3", report.Entries.Single(entry => entry.Term == "WordLoose").PageNumbersText);
                Assert.Equal("1 to 2", report.Entries.Single(entry => entry.Term == "WordRange").PageNumbersText);
                Assert.True(report.Entries.Single(entry => entry.Term == "WordSee").IsCrossReference);
                Assert.Equal("See WordLoose", report.Entries.Single(entry => entry.Term == "WordSee").CrossReferenceText);

                string indexText = TocText(index);
                Assert.Contains("Imported Separator Index", indexText);
                Assert.Contains("WordLoose => 2 | 3", indexText);
                Assert.Contains("WordRange => 1 to 2", indexText);
                Assert.Contains("WordSee -> See WordLoose", indexText);
                Assert.Contains(index.SdtBlock.Descendants<SimpleField>(), field => {
                    string instruction = field.Instruction?.Value ?? field.Instruction ?? string.Empty;
                    return instruction.Contains("\\e \" => \"", StringComparison.Ordinal) &&
                           instruction.Contains("\\l \" | \"", StringComparison.Ordinal) &&
                           instruction.Contains("\\g \" to \"", StringComparison.Ordinal) &&
                           instruction.Contains("\\k \" -> \"", StringComparison.Ordinal);
                });
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string indexText = TocText(index);

                Assert.Contains("WordLoose => 2 | 3", indexText);
                Assert.Contains("WordRange => 1 to 2", indexText);
                Assert.Contains("WordSee -> See WordLoose", indexText);
                Assert.Contains(index.SdtBlock.Descendants<SimpleField>(), field => {
                    string instruction = field.Instruction?.Value ?? field.Instruction ?? string.Empty;
                    return instruction.Contains("\\e \" => \"", StringComparison.Ordinal) &&
                           instruction.Contains("\\l \" | \"", StringComparison.Ordinal) &&
                           instruction.Contains("\\g \" to \"", StringComparison.Ordinal) &&
                           instruction.Contains("\\k \" -> \"", StringComparison.Ordinal);
                });
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexHonorsImportedEntryTypeFilters() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshEntryTypeFilter.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();
                SetIndexInstruction(index, " INDEX \\f \"A\" ");

                AddIndexEntryParagraph(document, "Alpha topic", " XE \"Alpha\" \\f \"A\" ");
                document.AddPageBreak();
                AddIndexEntryParagraph(document, "Beta topic", " XE \"Beta\" \\f \"B\" ");
                AddIndexEntryParagraph(document, "Loose topic", " XE \"Loose\" ");
                AddIndexEntryParagraph(document, "Gamma cross-reference", " XE \"Gamma\" \\f \"A\" \\t \"See Alpha\" ");

                WordIndexRefreshReport report = index.RefreshIndex("Typed Index");

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal(new[] { "Alpha", "Gamma" }, report.Entries.Select(entry => entry.Term).ToArray());
                Assert.All(report.Entries, entry => Assert.Equal("A", entry.EntryType));
                Assert.Equal(new[] { 1 }, report.Entries[0].PageNumbers);
                Assert.True(report.Entries[1].IsCrossReference);
                Assert.Equal("See Alpha", report.Entries[1].CrossReferenceText);

                string indexText = TocText(index);
                Assert.Contains("Typed Index", indexText);
                Assert.Contains("Alpha", indexText);
                Assert.Contains("Gamma", indexText);
                Assert.Contains("See Alpha", indexText);
                Assert.DoesNotContain("Beta", indexText);
                Assert.DoesNotContain("Loose", indexText);
                Assert.Contains(index.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\f \"A\"", StringComparison.Ordinal));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string indexText = TocText(index);

                Assert.Contains("Typed Index", indexText);
                Assert.Contains("Alpha", indexText);
                Assert.Contains("Gamma", indexText);
                Assert.DoesNotContain("Beta", indexText);
                Assert.DoesNotContain("Loose", indexText);
                Assert.Contains(index.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\f \"A\"", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexHonorsImportedBookmarkScopeFilters() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshBookmarkScope.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();
                SetIndexInstruction(index, " INDEX \\b \"ScopedIndex\" ");

                AddIndexEntryParagraph(document, "Outside before scope", " XE \"OutsideBefore\" ");

                string bookmarkId = document.BookmarkId.ToString();
                AppendBodyParagraph(document, new Paragraph(
                    new BookmarkStart { Name = "ScopedIndex", Id = bookmarkId },
                    new Run(new Text("Scoped alpha") { Space = SpaceProcessingModeValues.Preserve }),
                    CreateIndexEntryField(" XE \"ScopedAlpha\" ")));
                document.AddPageBreak();
                AppendBodyParagraph(document, new Paragraph(
                    new Run(new Text("Scoped beta") { Space = SpaceProcessingModeValues.Preserve }),
                    CreateIndexEntryField(" XE \"ScopedBeta\" "),
                    new BookmarkEnd { Id = bookmarkId }));

                AddIndexEntryParagraph(document, "Outside after scope", " XE \"OutsideAfter\" ");

                WordIndexRefreshReport report = index.RefreshIndex("Scoped Index");

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal(new[] { "ScopedAlpha", "ScopedBeta" }, report.Entries.Select(entry => entry.Term).ToArray());
                Assert.Equal(new[] { 1 }, report.Entries[0].PageNumbers);
                Assert.Equal(new[] { 2 }, report.Entries[1].PageNumbers);

                string indexText = TocText(index);
                Assert.Contains("Scoped Index", indexText);
                Assert.Contains("ScopedAlpha", indexText);
                Assert.Contains("ScopedBeta", indexText);
                Assert.DoesNotContain("OutsideBefore", indexText);
                Assert.DoesNotContain("OutsideAfter", indexText);
                Assert.Contains(index.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\b \"ScopedIndex\"", StringComparison.Ordinal));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string indexText = TocText(index);

                Assert.Contains("Scoped Index", indexText);
                Assert.Contains("ScopedAlpha", indexText);
                Assert.Contains("ScopedBeta", indexText);
                Assert.DoesNotContain("OutsideBefore", indexText);
                Assert.DoesNotContain("OutsideAfter", indexText);
                Assert.Contains(index.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\b \"ScopedIndex\"", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexSupportsBookmarkPageRanges() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshBookmarkPageRange.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();

                string bookmarkId = document.BookmarkId.ToString();
                AppendBodyParagraph(document, new Paragraph(
                    new BookmarkStart { Name = "RangeTopicBookmark", Id = bookmarkId },
                    new Run(new Text("Range topic starts") { Space = SpaceProcessingModeValues.Preserve })));
                document.AddPageBreak();
                AppendBodyParagraph(document, new Paragraph(
                    new Run(new Text("Range topic continues") { Space = SpaceProcessingModeValues.Preserve }),
                    new BookmarkEnd { Id = bookmarkId }));
                AddIndexEntryParagraph(document, "Range topic index entry", " XE \"Range Topic\" \\r \"RangeTopicBookmark\" ");
                AddIndexEntryParagraph(document, "Missing range topic", " XE \"Missing Range\" \\r \"MissingBookmark\" ");

                WordIndexRefreshReport report = index.RefreshIndex("Range Index");

                Assert.Equal(1, report.EntryCount);
                Assert.Equal(1, report.SkippedEntryCount);
                Assert.Equal("Range Topic", report.Entries[0].Term);
                Assert.Empty(report.Entries[0].PageNumbers);
                Assert.Equal(new[] { "1-2" }, report.Entries[0].PageReferences);
                Assert.Equal("1-2", report.Entries[0].PageNumbersText);

                string indexText = TocText(index);
                Assert.Contains("Range Index", indexText);
                Assert.Contains("Range Topic", indexText);
                Assert.Contains("1-2", indexText);
                Assert.DoesNotContain("Missing Range", indexText);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string indexText = TocText(index);

                Assert.Contains("Range Index", indexText);
                Assert.Contains("Range Topic", indexText);
                Assert.Contains("1-2", indexText);
                Assert.DoesNotContain("Missing Range", indexText);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexHonorsImportedCustomSeparators() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshCustomSeparators.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();
                SetIndexInstruction(index, " INDEX \\e \" => \" \\l \" | \" \\g \" to \" \\k \" -> \" ");

                AddIndexEntryParagraph(document, "Alpha topic", " XE \"Alpha\" ");
                document.AddPageBreak();
                AddIndexEntryParagraph(document, "Alpha follow-up", " XE \"Alpha\" ");

                string bookmarkId = document.BookmarkId.ToString();
                AppendBodyParagraph(document, new Paragraph(
                    new BookmarkStart { Name = "RangeTopicBookmark", Id = bookmarkId },
                    new Run(new Text("Range topic starts") { Space = SpaceProcessingModeValues.Preserve })));
                document.AddPageBreak();
                AppendBodyParagraph(document, new Paragraph(
                    new Run(new Text("Range topic continues") { Space = SpaceProcessingModeValues.Preserve }),
                    new BookmarkEnd { Id = bookmarkId }));
                AddIndexEntryParagraph(document, "Range topic index entry", " XE \"Range Topic\" \\r \"RangeTopicBookmark\" ");
                AddIndexEntryParagraph(document, "Gamma cross-reference", " XE \"Gamma\" \\t \"See Alpha\" ");

                WordIndexRefreshReport report = index.RefreshIndex("Separator Index");

                Assert.Equal(3, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal("1 | 2", report.Entries.Single(entry => entry.Term == "Alpha").PageNumbersText);
                Assert.Equal("2 to 3", report.Entries.Single(entry => entry.Term == "Range Topic").PageNumbersText);
                Assert.True(report.Entries.Single(entry => entry.Term == "Gamma").IsCrossReference);

                string indexText = TocText(index);
                Assert.Contains("Separator Index", indexText);
                Assert.Contains("Alpha => 1 | 2", indexText);
                Assert.Contains("Range Topic => 2 to 3", indexText);
                Assert.Contains("Gamma -> See Alpha", indexText);
                Assert.Contains(index.SdtBlock.Descendants<SimpleField>(), field => {
                    string instruction = field.Instruction?.Value ?? field.Instruction ?? string.Empty;
                    return instruction.Contains("\\e \" => \"", StringComparison.Ordinal) &&
                           instruction.Contains("\\l \" | \"", StringComparison.Ordinal) &&
                           instruction.Contains("\\g \" to \"", StringComparison.Ordinal) &&
                           instruction.Contains("\\k \" -> \"", StringComparison.Ordinal);
                });
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string indexText = TocText(index);

                Assert.Contains("Alpha => 1 | 2", indexText);
                Assert.Contains("Range Topic => 2 to 3", indexText);
                Assert.Contains("Gamma -> See Alpha", indexText);
                Assert.Contains(index.SdtBlock.Descendants<SimpleField>(), field => {
                    string instruction = field.Instruction?.Value ?? field.Instruction ?? string.Empty;
                    return instruction.Contains("\\e \" => \"", StringComparison.Ordinal) &&
                           instruction.Contains("\\l \" | \"", StringComparison.Ordinal) &&
                           instruction.Contains("\\g \" to \"", StringComparison.Ordinal) &&
                           instruction.Contains("\\k \" -> \"", StringComparison.Ordinal);
                });
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexHonorsImportedHeadingSeparators() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshHeadingSeparators.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();
                SetIndexInstruction(index, " INDEX \\h \"--A--\" ");

                AddIndexEntryParagraph(document, "Alpha topic", " XE \"Alpha\" ");
                AddIndexEntryParagraph(document, "Beta topic", " XE \"Beta\" ");
                AddIndexEntryParagraph(document, "Beta detail topic", " XE \"Beta:Detail\" ");
                AddIndexEntryParagraph(document, "Mango topic", " XE \"Mango\" ");
                AddIndexEntryParagraph(document, "Numeric topic", " XE \"1Number\" ");

                WordIndexRefreshReport report = index.RefreshIndex("Heading Separator Index");

                Assert.Equal(5, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal(new[] { "--A--", "--B--", "--M--" }, GetIndexHeadingTexts(index));

                string indexText = TocText(index);
                Assert.Contains("Heading Separator Index", indexText);
                Assert.Contains("--A--Alpha", indexText);
                Assert.Contains("--B--Beta", indexText);
                Assert.Contains("Detail", indexText);
                Assert.Contains("--M--Mango", indexText);
                Assert.Contains("1Number", indexText);
                Assert.Contains(index.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\h \"--A--\"", StringComparison.Ordinal));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Equal(new[] { "--A--", "--B--", "--M--" }, GetIndexHeadingTexts(index));
                Assert.Contains(index.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\h \"--A--\"", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexReportsAndPreservesImportedColumnCount() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshColumnCount.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();
                SetIndexInstruction(index, " INDEX \\c \"2\" ");

                AddIndexEntryParagraph(document, "Alpha topic", " XE \"Alpha\" ");
                AddIndexEntryParagraph(document, "Beta topic", " XE \"Beta\" ");

                WordIndexRefreshReport report = index.RefreshIndex("Two Column Index");

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal(2, report.ColumnCount);

                string indexText = TocText(index);
                Assert.Contains("Two Column Index", indexText);
                Assert.Contains("Alpha", indexText);
                Assert.Contains("Beta", indexText);
                Assert.Contains(index.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\c \"2\"", StringComparison.Ordinal));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains(index.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\c \"2\"", StringComparison.Ordinal));
                Assert.Contains("Alpha", TocText(index));
                Assert.Contains("Beta", TocText(index));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexCanUseConcordanceDocument() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshConcordance.docx");
            string concordancePath = Path.Combine(_directoryWithFiles, "IndexRefreshConcordanceSource.docx");
            File.Delete(filePath);
            File.Delete(concordancePath);

            using (WordDocument concordance = WordDocument.Create(concordancePath)) {
                AppendBodyElement(concordance, CreateSimpleTable(
                    CreateConcordanceRow("Alpha policy", "Policy:Alpha"),
                    CreateConcordanceRow("beta", "Evidence:Beta"),
                    CreateConcordanceRow("Gamma control", "Controls:Gamma"),
                    CreateConcordanceRow("Delta control", "Controls:Delta"),
                    CreateConcordanceRow("Epsilon text box", "TextBoxes:Epsilon"),
                    CreateConcordanceRow("ignored", "Unsafe \"Quote"),
                    CreateConcordanceRow(string.Empty, "Skipped")));
                concordance.Save(false);
            }

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();

                document.AddParagraph("Alpha policy uses beta evidence.");
                document.AddParagraph("alphabet soup should not match a partial concordance entry.");
                AppendBodyElement(document, CreateSimpleTable(
                    new TableRow(
                        new TableCell(
                            new Paragraph(
                                new Run(new Text("Gamma control owner") { Space = SpaceProcessingModeValues.Preserve }))))));
                AppendBodyElement(document, new SdtBlock(
                    new SdtProperties(
                        new SdtAlias { Val = "Concordance Content Control Source" },
                        new Tag { Val = "concordance-content-control-source" }),
                    new SdtContentBlock(
                        new Paragraph(
                            new Run(new Text("Delta control owner") { Space = SpaceProcessingModeValues.Preserve })))));
                AppendBodyParagraph(document, CreateTextBoxParagraph("Epsilon text box owner"));

                WordIndexConcordanceReport markReport = document.MarkIndexEntriesFromConcordance(concordancePath);

                Assert.Equal(5, markReport.ConcordanceEntryCount);
                Assert.Equal(5, markReport.MarkedEntryCount);
                Assert.Equal(4, markReport.MatchedParagraphCount);
                Assert.Equal(2, markReport.SkippedEntryCount);
                Assert.True(markReport.MatchWholeWord);
                Assert.False(markReport.MatchCase);

                string[] xeInstructions = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.XE)
                    .Select(field => field.InstructionText)
                    .ToArray();
                Assert.Contains(xeInstructions, instruction => instruction.Contains("Policy:Alpha", StringComparison.Ordinal));
                Assert.Contains(xeInstructions, instruction => instruction.Contains("Evidence:Beta", StringComparison.Ordinal));
                Assert.Contains(xeInstructions, instruction => instruction.Contains("Controls:Gamma", StringComparison.Ordinal));
                Assert.Contains(xeInstructions, instruction => instruction.Contains("Controls:Delta", StringComparison.Ordinal));
                Assert.Contains(xeInstructions, instruction => instruction.Contains("TextBoxes:Epsilon", StringComparison.Ordinal));
                Assert.DoesNotContain(xeInstructions, instruction => instruction.Contains("Unsafe", StringComparison.Ordinal));

                WordIndexRefreshReport indexReport = index.RefreshIndex("Concordance Index");

                Assert.Equal(5, indexReport.EntryCount);
                Assert.Equal(new[] { "Controls", "Controls", "Evidence", "Policy", "TextBoxes" }, indexReport.Entries.Select(entry => entry.Term).ToArray());
                Assert.Contains(indexReport.Entries, entry => entry.Term == "Policy" && entry.Subterm == "Alpha");
                Assert.Contains(indexReport.Entries, entry => entry.Term == "Evidence" && entry.Subterm == "Beta");
                Assert.Contains(indexReport.Entries, entry => entry.Term == "Controls" && entry.Subterm == "Gamma");
                Assert.Contains(indexReport.Entries, entry => entry.Term == "Controls" && entry.Subterm == "Delta");
                Assert.Contains(indexReport.Entries, entry => entry.Term == "TextBoxes" && entry.Subterm == "Epsilon");

                string indexText = TocText(index);
                Assert.Contains("Concordance Index", indexText);
                Assert.Contains("Policy", indexText);
                Assert.Contains("Alpha", indexText);
                Assert.Contains("Evidence", indexText);
                Assert.Contains("Beta", indexText);
                Assert.Contains("Controls", indexText);
                Assert.Contains("Gamma", indexText);
                Assert.Contains("Delta", indexText);
                Assert.Contains("TextBoxes", indexText);
                Assert.Contains("Epsilon", indexText);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains("Concordance Index", TocText(index));
                Assert.Equal(5, document.InspectFields().Count(field => field.FieldType == WordFieldType.XE));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexHonorsImportedLetterRangeFilters() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshLetterRange.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();
                SetIndexInstruction(index, " INDEX \\p \"A-M\" ");

                AddIndexEntryParagraph(document, "Alpha topic", " XE \"Alpha\" ");
                document.AddPageBreak();
                AddIndexEntryParagraph(document, "Mango topic", " XE \"Mango\" ");
                document.AddPageBreak();
                AddIndexEntryParagraph(document, "Zulu topic", " XE \"Zulu\" ");
                AddIndexEntryParagraph(document, "Numeric topic", " XE \"1Number\" ");

                WordIndexRefreshReport report = index.RefreshIndex("Letter Range Index");

                Assert.Equal(2, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal(new[] { "Alpha", "Mango" }, report.Entries.Select(entry => entry.Term).ToArray());

                string indexText = TocText(index);
                Assert.Contains("Letter Range Index", indexText);
                Assert.Contains("Alpha", indexText);
                Assert.Contains("Mango", indexText);
                Assert.DoesNotContain("Zulu", indexText);
                Assert.DoesNotContain("1Number", indexText);
                Assert.Contains(index.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\p \"A-M\"", StringComparison.Ordinal));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string indexText = TocText(index);

                Assert.Contains("Alpha", indexText);
                Assert.Contains("Mango", indexText);
                Assert.DoesNotContain("Zulu", indexText);
                Assert.DoesNotContain("1Number", indexText);
                Assert.Contains(index.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\p \"A-M\"", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexHonorsImportedLetterRangeFiltersWithSymbols() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshLetterRangeSymbols.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();
                SetIndexInstruction(index, " INDEX \\p \"!--B\" ");

                AddIndexEntryParagraph(document, "Symbol topic", " XE \"#Hash\" ");
                AddIndexEntryParagraph(document, "Alpha topic", " XE \"Alpha\" ");
                AddIndexEntryParagraph(document, "Beta topic", " XE \"Beta\" ");
                AddIndexEntryParagraph(document, "Cobalt topic", " XE \"Cobalt\" ");

                WordIndexRefreshReport report = index.RefreshIndex("Symbol Range Index");

                Assert.Equal(3, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal(new[] { "#Hash", "Alpha", "Beta" }, report.Entries.Select(entry => entry.Term).ToArray());

                string indexText = TocText(index);
                Assert.Contains("Symbol Range Index", indexText);
                Assert.Contains("#Hash", indexText);
                Assert.Contains("Alpha", indexText);
                Assert.Contains("Beta", indexText);
                Assert.DoesNotContain("Cobalt", indexText);
                Assert.Contains(index.SdtBlock.Descendants<SimpleField>(), field =>
                    (field.Instruction?.Value ?? field.Instruction ?? string.Empty).Contains("\\p \"!--B\"", StringComparison.Ordinal));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexReportsEmptyAndInvalidEntries() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshEmpty.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();

                AddIndexEntryParagraph(document, "Malformed index entry", " XE \"Alpha::Gamma\" ");

                WordIndexRefreshReport report = index.RefreshIndex("Generated Index");

                Assert.Equal(0, report.EntryCount);
                Assert.Equal(1, report.SkippedEntryCount);
                Assert.Contains("Generated Index", TocText(index));
                Assert.Contains("No index entries found.", TocText(index));
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexUsesNoteAnchorPageEstimates() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshEntriesNoteAnchorPages.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();

                AddIndexEntryParagraph(document, "Body topic", " XE \"BodyTopic\" ");
                document.AddPageBreak();
                document.AddParagraph("Footnote index anchor").AddFootNote("Footnote body");
                document.AddParagraph("Endnote index anchor").AddEndNote("Endnote body");
                Footnote footnote = document._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!.Elements<Footnote>().First(note => note.Type == null);
                Endnote endnote = document._wordprocessingDocument.MainDocumentPart!.EndnotesPart!.Endnotes!.Elements<Endnote>().First(note => note.Type == null);
                AppendIndexEntryParagraph(footnote, "Footnote topic", " XE \"FootnoteTopic\" ");
                AppendIndexEntryParagraph(endnote, "Endnote topic", " XE \"EndnoteTopic\" ");

                WordIndexRefreshReport report = index.RefreshIndex("Related-Part Index");

                Assert.Equal(new[] { "BodyTopic", "EndnoteTopic", "FootnoteTopic" }, report.Entries.Select(entry => entry.Term).ToArray());
                Assert.Equal(new[] { 1 }, report.Entries.Single(entry => entry.Term == "BodyTopic").PageNumbers);
                Assert.Equal(new[] { 2 }, report.Entries.Single(entry => entry.Term == "FootnoteTopic").PageNumbers);
                Assert.Equal(new[] { 2 }, report.Entries.Single(entry => entry.Term == "EndnoteTopic").PageNumbers);
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshIndexCombinesSplitXeFieldInstructions() {
            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshSplitXe.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTableOfContent index = document.AddTableOfContent();
                Paragraph paragraph = document.AddParagraph("Split topic")._paragraph;
                paragraph.Append(
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" XE \"Split") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FieldCode(" Topic\" ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text(string.Empty)),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

                WordIndexRefreshReport report = index.RefreshIndex("Split Index");

                WordIndexEntry entry = Assert.Single(report.Entries);
                Assert.Equal("Split Topic", entry.Term);
                Assert.Contains("Split Topic", TocText(index));
            }
        }

        private static void AssertGeneratedEntries(WordTableOfContent toc, params string[] expectedEntries) {
            string text = TocText(toc);
            foreach (string expectedEntry in expectedEntries) {
                Assert.Contains(expectedEntry, text);
            }
        }

        private static string TocText(WordTableOfContent toc) {
            return string.Concat(toc.SdtBlock.Descendants<Text>().Select(text => text.Text));
        }

        private static string[] GetIndexHeadingTexts(WordTableOfContent toc) {
            return toc.SdtBlock.Descendants<Paragraph>()
                .Where(paragraph => paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "IndexHeading")
                .Select(paragraph => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text)))
                .ToArray();
        }

        private static void AssertTocEntryPageNumberState(WordTableOfContent toc, string entryText, string expectedStyleId, bool shouldContainPageNumber) {
            Paragraph paragraph = Assert.Single(toc.SdtBlock.Descendants<Paragraph>(), item =>
                item.ParagraphProperties?.ParagraphStyleId?.Val == expectedStyleId &&
                string.Concat(item.Descendants<Text>().Select(text => text.Text)).Contains(entryText, StringComparison.Ordinal));

            Assert.Equal(shouldContainPageNumber, paragraph.Descendants<TabChar>().Any());
            Assert.Equal(shouldContainPageNumber, paragraph.Descendants<Text>().Any(text => text.Text == "1" || text.Text == "2"));
        }

        private static void AssertTocEntryPageNumberSeparator(WordTableOfContent toc, string entryText, string expectedStyleId, string expectedSeparator, string expectedPageNumber) {
            Paragraph paragraph = Assert.Single(toc.SdtBlock.Descendants<Paragraph>(), item =>
                item.ParagraphProperties?.ParagraphStyleId?.Val == expectedStyleId &&
                string.Concat(item.Descendants<Text>().Select(text => text.Text)).Contains(entryText, StringComparison.Ordinal));

            string paragraphText = string.Concat(paragraph.Descendants<Text>().Select(text => text.Text));
            Assert.Contains(entryText + expectedSeparator + expectedPageNumber, paragraphText);
            if (expectedSeparator != "\t") {
                Assert.Empty(paragraph.Descendants<TabChar>());
            }
        }

        private static void AddSectionBreakParagraph(WordDocument document, SectionMarkValues sectionMark) {
            Paragraph paragraph = new Paragraph(
                new ParagraphProperties(
                    new SectionProperties(
                        new SectionType {
                            Val = sectionMark
                        })));
            document._document.Body!.Append(paragraph);
        }

        private static void AddOutlineLevelParagraph(WordDocument document, string text, int outlineLevel) {
            WordParagraph paragraph = document.AddParagraph(text);
            paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph._paragraph.ParagraphProperties.OutlineLevel = new OutlineLevel { Val = outlineLevel };
        }

        private static void AddGeneratedCaptionParagraph(WordDocument document, string bookmarkName, string sequenceIdentifier, string captionText) {
            AddGeneratedCaptionParagraph(document, bookmarkName, sequenceIdentifier, "stale", captionText);
        }

        private static void AddGeneratedCaptionParagraph(WordDocument document, string bookmarkName, string sequenceIdentifier, string sequenceResult, string captionText) {
            string id = document.BookmarkId.ToString();

            Paragraph paragraph = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }),
                new BookmarkStart { Name = bookmarkName, Id = id },
                new Run(new Text(sequenceIdentifier + " ") { Space = SpaceProcessingModeValues.Preserve }),
                CreateCaptionSequenceField(sequenceIdentifier, sequenceResult),
                new Run(new Text(" " + captionText) { Space = SpaceProcessingModeValues.Preserve }),
                new BookmarkEnd { Id = id });

            AppendBodyParagraph(document, paragraph);
        }

        private static void AddTextBoxCaptionParagraph(WordDocument document, string sequenceIdentifier, string captionText) {
            Paragraph captionParagraph = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }),
                new Run(new Text(sequenceIdentifier + " ") { Space = SpaceProcessingModeValues.Preserve }),
                CreateCaptionSequenceField(sequenceIdentifier),
                new Run(new Text(" " + captionText) { Space = SpaceProcessingModeValues.Preserve }));

            Paragraph hostParagraph = new Paragraph(
                new Run(
                    new Picture(
                        new V.Shape(
                            new V.TextBox(
                                new TextBoxContent(captionParagraph))) {
                            Id = "OfficeIMO_TextBox_Caption",
                            Style = "width:240pt;height:40pt",
                            Filled = false,
                            Stroked = true
                        })));

            AppendBodyParagraph(document, hostParagraph);
        }

        private static Paragraph CreateParentAndTextBoxCaptionParagraph(string sequenceIdentifier, string parentCaptionText, string textBoxCaptionText) {
            Paragraph textBoxCaptionParagraph = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }),
                new Run(new Text(sequenceIdentifier + " ") { Space = SpaceProcessingModeValues.Preserve }),
                CreateCaptionSequenceField(sequenceIdentifier, "2"),
                new Run(new Text(" " + textBoxCaptionText) { Space = SpaceProcessingModeValues.Preserve }));

            return new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }),
                new Run(new Text(sequenceIdentifier + " ") { Space = SpaceProcessingModeValues.Preserve }),
                CreateCaptionSequenceField(sequenceIdentifier, "1"),
                new Run(new Text(" " + parentCaptionText) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(
                    new Picture(
                        new V.Shape(
                            new V.TextBox(
                                new TextBoxContent(textBoxCaptionParagraph))) {
                            Id = "OfficeIMO_TextBox_Anchored_Caption",
                            Style = "width:240pt;height:40pt",
                            Filled = false,
                            Stroked = true
                        })));
        }

        private static SimpleField CreateCaptionSequenceField(string sequenceIdentifier) {
            return CreateCaptionSequenceField(sequenceIdentifier, "stale");
        }

        private static SimpleField CreateCaptionSequenceField(string sequenceIdentifier, string resultText) {
            return new SimpleField(
                new Run(
                    new Text(resultText) { Space = SpaceProcessingModeValues.Preserve })) {
                Instruction = " SEQ " + sequenceIdentifier + " "
            };
        }

        private static void AppendCaptionParagraph(OpenXmlCompositeElement root, string sequenceIdentifier, string sequenceResult, string captionText) {
            root.Append(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }),
                new Run(new Text(sequenceIdentifier + " ") { Space = SpaceProcessingModeValues.Preserve }),
                CreateCaptionSequenceField(sequenceIdentifier, sequenceResult),
                new Run(new Text(" " + captionText) { Space = SpaceProcessingModeValues.Preserve })));
        }

        private static void AddIndexEntryParagraph(WordDocument document, string visibleText, string instruction) {
            AppendBodyParagraph(document, CreateIndexSourceParagraph(visibleText, instruction));
        }

        private static void AppendIndexEntryParagraph(OpenXmlCompositeElement root, string visibleText, string instruction) {
            root.Append(CreateIndexSourceParagraph(visibleText, instruction));
        }

        private static Paragraph CreateIndexSourceParagraph(string visibleText, string instruction) {
            Paragraph paragraph = new Paragraph(
                new Run(new Text(visibleText) { Space = SpaceProcessingModeValues.Preserve }),
                CreateIndexEntryField(instruction));

            return paragraph;
        }

        private static SimpleField CreateIndexEntryField(string instruction) {
            return new SimpleField(new Run(new Text(string.Empty))) {
                Instruction = instruction
            };
        }

        private static Table CreateSimpleTable(params TableRow[] rows) {
            int columnCount = Math.Max(1, rows.Select(row => row.Elements<TableCell>().Count()).DefaultIfEmpty(1).Max());
            Table table = new Table(
                new TableProperties(
                    new TableWidth {
                        Width = "0",
                        Type = TableWidthUnitValues.Auto
                    }),
                new TableGrid(Enumerable.Range(0, columnCount).Select(_ => new GridColumn()).Cast<OpenXmlElement>()));
            table.Append(rows);
            return table;
        }

        private static TableRow CreateConcordanceRow(string searchText, string indexText) {
            return new TableRow(
                CreateConcordanceCell(searchText),
                CreateConcordanceCell(indexText));
        }

        private static TableCell CreateConcordanceCell(string text) {
            return new TableCell(
                new Paragraph(
                    new Run(
                        new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
        }

        private static Paragraph CreateTextBoxParagraph(string text) {
            return new Paragraph(
                new Run(
                    new Picture(
                        new V.Shape(
                            new V.TextBox(
                                new TextBoxContent(
                                    new Paragraph(
                                        new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }))))) {
                            Id = "OfficeIMO_Concordance_TextBox",
                            Style = "width:240pt;height:40pt",
                            Filled = false,
                            Stroked = true
                        })));
        }

        private static Paragraph CreateTextBoxHeadingParagraph(string text, string styleId) {
            return new Paragraph(
                new Run(
                    new Picture(
                        new V.Shape(
                            new V.TextBox(
                                new TextBoxContent(
                                    new Paragraph(
                                        new ParagraphProperties(new ParagraphStyleId { Val = styleId }),
                                        new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }))))) {
                            Id = "OfficeIMO_Toc_TextBox",
                            Style = "width:240pt;height:40pt",
                            Filled = false,
                            Stroked = true
                        })));
        }

        private static Paragraph CreateParentAndTextBoxHeadingParagraph(string parentText, string textBoxText) {
            return new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                new Run(new Text(parentText) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(
                    new Picture(
                        new V.Shape(
                            new V.TextBox(
                                new TextBoxContent(
                                    new Paragraph(
                                        new ParagraphProperties(new ParagraphStyleId { Val = "Heading2" }),
                                        new Run(new Text(textBoxText) { Space = SpaceProcessingModeValues.Preserve }))))) {
                            Id = "OfficeIMO_Toc_Parent_TextBox",
                            Style = "width:240pt;height:40pt",
                            Filled = false,
                            Stroked = true
                        })));
        }

        private static void SetIndexInstruction(WordTableOfContent index, string instruction) {
            SimpleField field = index.SdtBlock.Descendants<SimpleField>().First();
            field.Instruction = instruction;
        }

        private static void SetTocInstruction(WordTableOfContent toc, string instruction) {
            SimpleField field = toc.SdtBlock.Descendants<SimpleField>().First();
            field.Instruction = instruction;
        }

        private static Paragraph CreateComplexFieldParagraph(string instruction, string resultText) {
            return new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(resultText) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        private static void AppendBodyParagraph(WordDocument document, Paragraph paragraph) {
            AppendBodyElement(document, paragraph);
        }

        private static void AppendBodyElement(WordDocument document, OpenXmlElement element) {
            Body body = document._document.Body!;
            SectionProperties? finalSectionProperties = body.Elements<SectionProperties>().LastOrDefault();
            if (finalSectionProperties != null) {
                body.InsertBefore(element, finalSectionProperties);
            } else {
                body.Append(element);
            }
        }
    }
}
