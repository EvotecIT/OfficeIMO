using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void Html_MixedDirections_ArePreserved() {
            string html = "<p dir='ltr'>Left</p><p dir='rtl'>يمين</p><p style='direction:rtl'>CSS</p><p><bdo dir='ltr'>Override</bdo></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            Assert.False(doc.Paragraphs[0].BiDi);
            Assert.True(doc.Paragraphs[1].BiDi);
            Assert.True(doc.Paragraphs[2].BiDi);
            Assert.False(doc.Paragraphs[3].BiDi);
        }
    }
}

