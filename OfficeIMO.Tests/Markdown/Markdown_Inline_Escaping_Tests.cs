using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Inline_Escaping_Tests {
        [Fact]
        public void TextRun_EscapesReservedCharacters() {
            var run = new TextRun("[demo](test)|back\\slash");

            var markdown = run.RenderMarkdown();

            Assert.Equal("\\[demo\\]\\(test\\)\\|back\\\\slash", markdown);
        }

        [Fact]
        public void DecodedHtmlEntityTextRun_EncodesLiteralAngleBrackets() {
            IRenderableMarkdownInline run = new DecodedHtmlEntityTextRun("<u>demo</u>");

            var markdown = run.RenderMarkdown();

            Assert.Equal("&lt;u&gt;demo&lt;/u&gt;", markdown);
        }

        [Fact]
        public void DecodedHtmlEntityTextRun_EscapesLiteralMarkdownDelimiters() {
            IRenderableMarkdownInline run = new DecodedHtmlEntityTextRun("`code` ~~strike~~ ==mark==");

            var markdown = run.RenderMarkdown();

            Assert.Equal(@"\`code\` \~\~strike\~\~ \=\=mark\=\=", markdown);
        }

        [Fact]
        public void CommonMark_Profile_Decodes_Html5_Named_Character_References() {
            const string markdown = "&nbsp; &amp; &copy; &AElig; &Dcaron;\n&frac34; &HilbertSpace; &DifferentialD;\n&ClockwiseContourIntegral; &CounterClockwiseContourIntegral; &ngE;\n";
            const string expected = "<p>  &amp; © Æ Ď\n¾ ℋ ⅆ\n∲ ∳ ≧̸</p>\n";

            var document = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
            var html = document.ToHtmlFragment(CommonMarkHtmlComparison.CreatePlainHtmlOptions());

            Assert.Equal(CommonMarkHtmlComparison.Normalize(expected), CommonMarkHtmlComparison.Normalize(html));
        }

        [Fact]
        public void CommonMark_Profile_Decodes_Numeric_Character_References_With_Replacement() {
            const string markdown = "&#35; &#1234; &#992; &#0;\n";
            const string expected = "<p># Ӓ Ϡ �</p>\n";

            var document = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
            var html = document.ToHtmlFragment(CommonMarkHtmlComparison.CreatePlainHtmlOptions());

            Assert.Equal(CommonMarkHtmlComparison.Normalize(expected), CommonMarkHtmlComparison.Normalize(html));
        }

        [Fact]
        public void LinkAndImageInline_EscapeTextAndUrls() {
            var link = new LinkInline("[text]", "path(1)|two\\end", "see [ref] | note");
            var image = new ImageInline("alt[text]", "image(path)|pipe\\end", "badge [info] | note");

            Assert.Equal("[\\[text\\]](path\\(1\\)\\|two\\\\end \"see \\[ref\\] \\| note\")", link.RenderMarkdown());
            Assert.Equal("![alt\\[text\\]](image\\(path\\)\\|pipe\\\\end \"badge \\[info\\] \\| note\")", image.RenderMarkdown());
        }

        [Fact]
        public void Emphasis_EscapesReservedCharacters() {
            var bold = new BoldInline("[bold] (text) | back\\slash");
            var strike = new StrikethroughInline("[strike](target)|pipe\\");

            Assert.Equal("**\\[bold\\] \\(text\\) \\| back\\\\slash**", bold.RenderMarkdown());
            Assert.Equal("~~\\[strike\\]\\(target\\)\\|pipe\\\\~~", strike.RenderMarkdown());
        }
    }
}
