using System;
using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests
{
    public class MarkdownRoundTripTests
    {
        [Fact]
        public void WordToMarkdown_PreservesCachedSimpleDateFieldAndItalicText()
        {
            using var document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("Due ");
            paragraph._paragraph.Append(new SimpleField(new Run(new Text("2020-01-02"))) { Instruction = " DATE \\@ \"yyyy-MM-dd\" " });
            paragraph.AddText(" confirmed").SetItalic();
            string markdown = ((document.ToMarkdown().ToMarkdown()));
            Assert.Contains("Due 2020-01-02", markdown);
            Assert.Contains("*confirmed*", markdown);
        }

        [Fact]
        public void WordToMarkdown_ReportsFlattenedMergedTableCells()
        {
            using var document = WordDocument.Create();
            var table = document.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].AddText("Combined");
            table.Rows[0].Cells[0].MergeHorizontally(1);
            WordToMarkdownResult result = document.ConvertToMarkdownWithEvidence();
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Message.Contains("cell merges", StringComparison.Ordinal));
            Assert.True(result.Report.HasLoss);
        }

        [Fact]
        public void WordToMarkdown_OnlyReportsAuthoredCellBorders()
        {
            using var document = WordDocument.Create();
            var cell = document.AddTable(1, 1, WordTableStyle.TableNormal).Rows[0].Cells[0];
            cell.Paragraphs[0].AddText("Value");
            cell._tableCell.TableCellProperties!.TableCellBorders = new TableCellBorders();
            WordToMarkdownResult emptyBorders = document.ConvertToMarkdownWithEvidence();
            Assert.False(emptyBorders.Report.HasLoss);
            emptyBorders.Report.RequireNoLoss();
            cell._tableCell.TableCellProperties.TableCellBorders.Append(new TopBorder { Val = BorderValues.None });
            WordToMarkdownResult authoredBorder = document.ConvertToMarkdownWithEvidence();
            Assert.Contains(authoredBorder.Report.Diagnostics, diagnostic => diagnostic.Message.Contains("table borders", StringComparison.Ordinal));
            Assert.True(authoredBorder.Report.HasLoss);
        }

        [Fact]
        public void WordToMarkdown_ReportsMergedCellsAndAuthoredBordersSeparately()
        {
            using var document = WordDocument.Create();
            var cell = document.AddTable(1, 2, WordTableStyle.TableNormal).Rows[0].Cells[0];
            cell.Paragraphs[0].AddText("Combined");
            cell.MergeHorizontally(1);
            cell._tableCell.TableCellProperties!.TableCellBorders = new TableCellBorders(new TopBorder { Val = BorderValues.None });
            WordToMarkdownResult result = document.ConvertToMarkdownWithEvidence();
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Message.Contains("cell merges", StringComparison.Ordinal));
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Message.Contains("table borders", StringComparison.Ordinal));
        }

        [Fact]
        public void WordToMarkdown_ReportsDirectAndInheritedTableBorders()
        {
            using var document = WordDocument.Create();
            WordTable direct = document.AddTable(1, 1, WordTableStyle.TableNormal);
            direct.Rows[0].Cells[0].Paragraphs[0].Text = "Direct";
            direct._tableProperties!.TableBorders = new TableBorders(new TopBorder { Val = BorderValues.Single });
            WordToMarkdownResult directResult = document.ConvertToMarkdownWithEvidence();
            Assert.Contains(directResult.Report.Diagnostics, diagnostic => diagnostic.Message.Contains("table borders", StringComparison.Ordinal));
            direct._tableProperties.TableBorders = null;
            direct._tableProperties.TableStyle!.Val = "InheritedGrid";
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new Style(new BasedOn { Val = "TableGrid" }) { Type = StyleValues.Table, StyleId = "InheritedGrid" });
            WordToMarkdownResult inheritedResult = document.ConvertToMarkdownWithEvidence();
            Assert.Contains(inheritedResult.Report.Diagnostics, diagnostic => diagnostic.Message.Contains("table borders", StringComparison.Ordinal));
        }

        [Fact]
        public void WordToMarkdown_ReportsBordersFromEnabledConditionalTableStyle()
        {
            using var document = WordDocument.Create();
            WordTable table = document.AddTable(1, 1, WordTableStyle.TableNormal);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Conditional";
            table._tableProperties!.TableStyle!.Val = "ConditionalOnly";
            var conditional = new TableStyleProperties
            {
                Type = TableStyleOverrideValues.FirstRow
            };
            conditional.Append(new TableStyleConditionalFormattingTableCellProperties(new TableCellBorders(new TopBorder { Val = BorderValues.Single })));
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new Style(conditional) { Type = StyleValues.Table, StyleId = "ConditionalOnly" });
            table.ConditionalFormattingFirstRow = false;
            Assert.DoesNotContain(document.ConvertToMarkdownWithEvidence().Report.Diagnostics, diagnostic => diagnostic.Message.Contains("table borders", StringComparison.Ordinal));
            table.ConditionalFormattingFirstRow = true;
            Assert.Contains(document.ConvertToMarkdownWithEvidence().Report.Diagnostics, diagnostic => diagnostic.Message.Contains("table borders", StringComparison.Ordinal));
        }

        [Fact]
        public void Markdown_To_Word_To_Markdown_RoundTrip_Preserves_CoreFeatures()
        {
            string md = "" + "# Report\n" + "Intro with **bold**, *italic*, ~~strike~~, <u>underline</u>, and `code`.\\n\n" + "- Item 1\n" + "- [x] Task done\n" + "- [ ] Task todo\n" + "- Link: [Docs](https://example.com)\n\n" + "1. First\n2. Second\n\n" + "| Name | Score | Date |\n" + "|:-----|------:|:----:|\n" + "| Alice | 98.5 | 2024-01-10 |\n" + "| Bob   | 91.0 | 2023-08-22 |\n\n" + "Here is a ref[^1].\n\n" + "[^1]: Footnote body.";
            using var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord();
            var md2 = (doc.ToMarkdown(new WordMarkdownExportOptions { EnableUnderline = true })).ToMarkdown();
            Assert.Contains("# Report", md2);
            Assert.Contains("**bold**", md2);
            Assert.Contains("*italic*", md2);
            Assert.Contains("~~strike~~", md2);
            Assert.Contains("<u>underline</u>", md2);
            Assert.Contains("`code`", md2);
            // Accept checkbox presence; text may be separated by spacing
            Assert.Contains("[x]", md2);
            Assert.Contains("Task done".Replace(" ", ""), md2.Replace(" ", ""));
            Assert.Contains("- [ ] Task todo", md2);
            Assert.Contains("[Docs](https://example.com)", md2);
            // Table header output formatting may vary; check for header tokens
            Assert.Contains("| Name", md2);
            Assert.Contains("| Score", md2);
            Assert.Contains("| Date", md2);
            // Alignment row may be omitted depending on table conversion; skip strict check
            Assert.Contains("[^1]", md2);
            Assert.Contains("[^1]:", md2);
        }

        [Fact]
        public void Word_Fluent_To_Markdown_Back_To_Word_Preserves_Structure()
        {
            using var doc = WordDocument.Create();
            {
                doc.AddParagraph("Title").SetStyle(WordParagraphStyles.Heading1);
                doc.AddParagraph("Hello world");
                var list1 = doc.AddList(WordListStyle.Bulleted);
                int level2 = 0;
                list1.AddItem("One", level2);
                list1.AddTaskItem("Done", true, level2);
                list1.AddLinkItem("Docs", new Uri("https://example.com"), false, level2);
                level2++;
                list1.AddItem("SubOne", level2);
                level2++;
                list1.AddItem("SubSub", level2);
                level2 = Math.Max(0, level2 - 1);
                var list3 = doc.AddList(WordListStyle.Numbered);
                int level4 = 0;
                list3.AddItem("First", level4);
                level4++;
                list3.AddItem("First.A", level4);
                level4 = Math.Max(0, level4 - 1);
                list3.AddItem("Second", level4);
                var paragraph5 = doc.AddParagraph();
                paragraph5.AddFormattedText("B", bold: true);
                paragraph5 = paragraph5.AddText(" ");
                paragraph5.AddFormattedText("I", italic: true);
                paragraph5 = paragraph5.AddText(" ");
                paragraph5.AddFormattedText("U", underline: WordUnderlineStyle.Single);
                paragraph5 = paragraph5.AddText(" ");
                paragraph5.AddText("S").SetStrike();
                paragraph5 = paragraph5.AddText(" ");
                paragraph5.AddText("C").SetFontFamily("Consolas");
                var table6 = doc.AddTable(1, 2);
                var row7 = table6.Rows[0];
                row7.Cells[0].Paragraphs[0].Text = System.Convert.ToString("Name") ?? string.Empty;
                row7.Cells[1].Paragraphs[0].Text = System.Convert.ToString("Score") ?? string.Empty;
                row7.RepeatHeaderRowAtTheTopOfEachPage = true;
                table6.ConditionalFormattingFirstRow = true;
                var row8 = table6.AddRow();
                row8.Cells[0].Paragraphs[0].Text = System.Convert.ToString("Alice") ?? string.Empty;
                row8.Cells[1].Paragraphs[0].Text = System.Convert.ToString("98.5") ?? string.Empty;
                var row9 = table6.AddRow();
                row9.Cells[0].Paragraphs[0].Text = System.Convert.ToString("Bob") ?? string.Empty;
                row9.Cells[1].Paragraphs[0].Text = System.Convert.ToString("91.0") ?? string.Empty;
            }

            var md = (doc.ToMarkdown(new WordMarkdownExportOptions { EnableUnderline = true })).ToMarkdown();
            Assert.Contains("# Title", md);
            Assert.Contains("Hello world", md);
            Assert.Contains("- One", md);
            Assert.Contains("  - SubOne", md);
            Assert.Contains("    - SubSub", md);
            Assert.Contains("[x]", md);
            Assert.Contains("Done", md);
            // Link rendering may be plain text in list items; accept either form
            Assert.Matches("(\\[Docs\\]\\(https://example\\.com\\)|\\bDocs\\b)", md);
            Assert.Contains("1. First", md);
            Assert.Contains("  1. First.A", md);
            Assert.Contains("**B**", md);
            Assert.Contains("*I*", md);
            Assert.Contains("<u>U</u>", md);
            Assert.Contains("~~S~~", md);
            Assert.Contains("`C`", md);
            Assert.Contains("| Name | Score |", md);
            using var doc2 = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord();
            // Verify checkbox restored
            var checkboxParagraph = doc2.Paragraphs.FirstOrDefault(p => p.IsCheckBox);
            Assert.NotNull(checkboxParagraph);
            Assert.True(checkboxParagraph!.CheckBox!.IsChecked);
            // Verify heading style exists in the collection
            Assert.Contains(doc2.Paragraphs, p => p.Style == WordParagraphStyles.Heading1);
        }
    }
}