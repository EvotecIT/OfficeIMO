using AngleSharp.Html.Parser;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordHtmlStyles_KeepParagraphWrappedListItemsCompactAndHonorExplicitSpacing(bool sharedShell) {
        using WordDocument word = WordDocument.Create();
        string exported = word.ToHtml(new WordHtmlExportOptions { UseSharedDocumentShell = sharedShell });
        var parsed = new HtmlParser().ParseDocument(exported);
        string styles = string.Join("", parsed.QuerySelectorAll("style").Select(style => style.OuterHtml));
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            ViewportWidth = 600D,
            ViewportHeight = 1000D,
            Margins = HtmlRenderMargins.All(0D)
        };

        double ItemDistance(string items) {
            HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(styles + "<ul>" + items + "</ul>", options);
            HtmlRenderText[] text = EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>().ToArray();
            return Assert.Single(text, item => item.Text == "Second").Y
                - Assert.Single(text, item => item.Text == "First").Y;
        }

        double plain = ItemDistance("<li>First</li><li>Second</li>");
        double wrapped = ItemDistance("<li><p>First</p></li><li><p>Second</p></li>");
        double authored = ItemDistance("<li><p style='margin-bottom:12px'>First</p></li><li><p>Second</p></li>");

        Assert.Equal(plain, wrapped, precision: 3);
        Assert.True(authored > wrapped + 1D, "An explicit paragraph margin must still add space between list items.");
    }
}
