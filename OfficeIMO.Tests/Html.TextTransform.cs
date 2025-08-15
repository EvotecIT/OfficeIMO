using System.Linq;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Theory]
        [InlineData("uppercase", "HeLLo wORld", "HELLO WORLD")]
        [InlineData("lowercase", "HeLLo wORld", "hello world")]
        [InlineData("capitalize", "hello world", "Hello World")]
        public void HtmlToWord_TextTransform(string transform, string input, string expected) {
            string html = $"<p style=\"text-transform:{transform}\">{input}</p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var run = doc.Paragraphs.First();
            Assert.Equal(expected, run.Text);
        }
    }
}
