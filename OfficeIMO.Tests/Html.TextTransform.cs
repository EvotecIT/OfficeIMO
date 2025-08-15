using System.Linq;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Theory]
        [InlineData("uppercase", "Hello World", "HELLO WORLD")]
        [InlineData("lowercase", "Hello World", "hello world")]
        [InlineData("capitalize", "hello world", "Hello World")]
        public void HtmlToWord_TextTransform(string transform, string input, string expected) {
            string html = $"<p><span style=\"text-transform:{transform}\">{input}</span></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var run = doc.Paragraphs[0].GetRuns().First();

            Assert.Equal(expected, run.Text);
        }
    }
}
