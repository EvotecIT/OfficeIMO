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
    }
}

