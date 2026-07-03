using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void PreDefaultRendersParagraphs() {
            string html = "<pre><code>var x = 1;\nvar y = 2;</code></pre>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            Assert.Empty(doc.Tables);
            var codeParas = doc.Paragraphs.Where(p => p.StyleId == "HTMLPreformatted" && !string.IsNullOrEmpty(p.Text)).ToList();
            Assert.Equal(2, codeParas.Count);
        }

        [Fact]
        public void PreAsTableRendersTable() {
            string html = "<pre><code>var x = 1;\nvar y = 2;</code></pre>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions { RenderPreAsTable = true });

            Assert.Single(doc.Tables);
            var table = doc.Tables[0];
            Assert.Equal(1, table.RowsCount);
            Assert.Equal(1, table.Rows[0].CellsCount);
            var cell = table.Rows[0].Cells[0];
            var paras = cell.Paragraphs.Where(p => p.StyleId == "HTMLPreformatted" && !string.IsNullOrEmpty(p.Text)).ToList();
            Assert.Equal(2, paras.Count);
            Assert.Equal("var x = 1;", paras[0].Text);
            Assert.Equal("var y = 2;", paras[1].Text);
        }
    }
}
