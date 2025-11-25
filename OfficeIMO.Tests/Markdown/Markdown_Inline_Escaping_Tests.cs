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
