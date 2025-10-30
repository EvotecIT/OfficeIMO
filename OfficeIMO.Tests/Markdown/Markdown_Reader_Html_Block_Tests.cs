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
        public void Type7_Tag_Requiring_Closing_Bracket_Remains_Text_When_Incomplete() {
            string md = "<span class=\"note\"\nParagraph";

            var doc = MarkdownReader.Parse(md);

            Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            Assert.DoesNotContain(doc.Blocks, block => block is HtmlRawBlock);
        }
    }
}

