using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_WhiteSpace_Normal() {
            string html = "<p style=\"white-space:normal\">Hello   world\nFoo</p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var text = string.Concat(doc.Paragraphs[0].GetRuns().Where(r => !r.IsBreak).Select(r => r.Text));
            Assert.Equal("Hello world Foo", text);
        }

        [Fact]
        public void HtmlToWord_WhiteSpace_Pre() {
            string html = "<p style=\"white-space:pre\">Hello   world\nFoo</p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var runs = doc.Paragraphs[0].GetRuns().Where(r => !r.IsBreak).ToArray();
            Assert.Equal("Hello\u00A0\u00A0\u00A0world", runs[0].Text);
            Assert.Equal("Foo", runs[1].Text);
        }

        [Fact]
        public void HtmlToWord_WhiteSpace_PreWrap() {
            string html = "<p style=\"white-space:pre-wrap\">Hello   world\nFoo</p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var runs = doc.Paragraphs[0].GetRuns().Where(r => !r.IsBreak).ToArray();
            Assert.Equal("Hello   world", runs[0].Text);
            Assert.Equal("Foo", runs[1].Text);
        }

        [Fact]
        public void HtmlToWord_WhiteSpace_NoWrap() {
            string html = "<p style=\"white-space:nowrap\">Hello   world\nFoo</p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var text = string.Concat(doc.Paragraphs[0].GetRuns().Where(r => !r.IsBreak).Select(r => r.Text));
            Assert.Equal("Hello\u00A0world\u00A0Foo", text);
        }
    }
}
