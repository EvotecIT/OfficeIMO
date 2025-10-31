using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Reader_Html_Block_Tests {
        [Fact]
        public void Parses_Gfm_Type1_Script_Block() {
            string md = "<script>\nlet x = 1;\n</script>\n\nParagraph.";

            var doc = MarkdownReader.Parse(md);

            Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<script>\nlet x = 1;\n</script>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Parses_Gfm_Type2_Html_Comment_Block() {
            string md = "<!-- start\ncontinues -->\n\nNext";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<!-- start\ncontinues -->", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Parses_Gfm_Type6_Block_Until_Blank_Line() {
            string md = "<div>\n<p>inline</p>\n\nparagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<div>\n<p>inline</p>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Type6_Closing_Tag_Starts_Block() {
            string md = "</div>\n\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("</div>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Parses_Gfm_Type3_Processing_Instruction_Block() {
            string md = "<?xml version=\"1.0\"?>\n\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<?xml version=\"1.0\"?>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Parses_Gfm_Type4_Declaration_Block() {
            string md = "<!DOCTYPE html>\n\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<!DOCTYPE html>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Parses_Gfm_Type5_CData_Block() {
            string md = "<![CDATA[<p>literal</p>]]>\n\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<![CDATA[<p>literal</p>]]>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Parses_Gfm_Type7_Generic_Tag_Block() {
            string md = "<span class=\"note\">\ninline\n</span>\n\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<span class=\"note\">\ninline\n</span>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Incomplete_Block_Tag_Remains_Text() {
            string md = "<div\nParagraph";

            var doc = MarkdownReader.Parse(md);

            Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            Assert.DoesNotContain(doc.Blocks, block => block is HtmlRawBlock);
        }

        [Fact]
        public void Script_Start_Requires_Terminating_Angle_Bracket() {
            string md = "<script\nalert(1);";

            var doc = MarkdownReader.Parse(md);

            Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
        }

        [Fact]
        public void Type6_Tag_With_Escaped_Quotes_Is_Recognized() {
            string md = "<div data-json=\"{\\\"key\\\":\\\"value\\\"}\">\ncontent\n\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<div data-json=\"{\\\"key\\\":\\\"value\\\"}\">\ncontent", html.Html);
        }

        [Fact]
        public void Type6_Block_Ends_When_Stack_Unwinds() {
            string md = "<div>\n<section>\n<p>Value</p>\n</section>\n</div>\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<div>\n<section>\n<p>Value</p>\n</section>\n</div>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Type7_Tag_Requiring_Closing_Bracket_Remains_Text_When_Incomplete() {
            string md = "<span class=\"note\"\nParagraph";

            var doc = MarkdownReader.Parse(md);

            Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            Assert.DoesNotContain(doc.Blocks, block => block is HtmlRawBlock);
        }

        [Fact]
        public void Type6_Details_Block_Preserves_Blank_Line_Content() {
            string md = "<details>\n<summary>Summary</summary>\n\n<div>Body</div>\n</details>\n\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<details>\n<summary>Summary</summary>\n\n<div>Body</div>\n</details>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Type6_Table_Block_Preserves_Blank_Line_Separated_Sections() {
            string md = "<table>\n<thead>\n<tr><th>H</th></tr>\n</thead>\n\n<tbody>\n<tr><td>R1</td></tr>\n</tbody>\n</table>";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<table>\n<thead>\n<tr><th>H</th></tr>\n</thead>\n\n<tbody>\n<tr><td>R1</td></tr>\n</tbody>\n</table>", html.Html);
        }

        [Fact]
        public void Type6_Closing_Tag_Does_Not_Consume_Following_Html() {
            string md = "</div>\n\n<div>\n<p>Next</p>\n</div>";

            var doc = MarkdownReader.Parse(md);

            var closingBlock = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("</div>", closingBlock.Html);

            var nextBlock = Assert.IsType<HtmlRawBlock>(doc.Blocks[1]);
            Assert.Equal("<div>\n<p>Next</p>\n</div>", nextBlock.Html);
        }

        [Fact]
        public void Type6_Details_Block_Allows_Blank_Line_Before_Closing_Tag() {
            string md = "<details>\n<summary>One</summary>\n\n<section>\n<p>Inner</p>\n</section>\n\n</details>\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<details>\n<summary>One</summary>\n\n<section>\n<p>Inner</p>\n</section>\n\n</details>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Type6_Details_Block_With_Nested_Table_And_Blank_Lines_Remains_Intact() {
            string md = "<details>\n<summary>Summary</summary>\n\n<table>\n<thead>\n<tr><th>H</th></tr>\n</thead>\n\n<tbody>\n<tr><td>R1</td></tr>\n</tbody>\n</table>\n\n<div>Tail</div>\n</details>";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<details>\n<summary>Summary</summary>\n\n<table>\n<thead>\n<tr><th>H</th></tr>\n</thead>\n\n<tbody>\n<tr><td>R1</td></tr>\n</tbody>\n</table>\n\n<div>Tail</div>\n</details>", html.Html);
        }

        [Fact]
        public void Type6_Details_Block_Allows_SelfClosing_Child_Elements() {
            string md = "<details>\n<summary>Summary</summary>\n<component />\n\n<div>Body</div>\n</details>\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<details>\n<summary>Summary</summary>\n<component />\n\n<div>Body</div>\n</details>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }

        [Fact]
        public void Type6_Details_Block_Closes_With_Unmatched_Inner_Tags() {
            string md = "<details>\n<div>\n<p>Loose</p>\n</details>\nParagraph";

            var doc = MarkdownReader.Parse(md);

            var html = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
            Assert.Equal("<details>\n<div>\n<p>Loose</p>\n</details>", html.Html);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        }
    }
}

