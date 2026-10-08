using OfficeIMO.Epub;
using OfficeIMO.Html;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubCssRewritePreservationContracts {
    [Fact]
    public void UrlRepairPreservesCompoundSelectorsCommentsAndUnrelatedCarriers() {
        const string source = "/* url(ignored.png) */\r\n@import 'unchanged.css' print;\r\n" +
            "#heading/**/.lead{color:#13579b;background:url('old.png');content:'/* literal */'}\r\n" +
            "p { background-image:image-set('untouched.png' 1x, 'old.png' 2x) } /* trailing */";
        var visited = new List<string>();
        string result = HtmlResourcePipeline.RewriteCssResourceUrls(source, (url, _) => {
            visited.Add(url); return url == "old.png" ? "new.png" : url;
        });
        Assert.Equal(source.Replace("url('old.png')", "url(\"new.png\")").Replace("'old.png' 2x", "\"new.png\" 2x"), result);
        Assert.DoesNotContain("ignored.png", visited);
        Assert.Equal(source, HtmlResourcePipeline.RewriteCssResourceUrls(source, (url, _) => url));
    }

    [Fact]
    public void ImportReplacementRetainsSurroundingCommentsAndConditions() {
        const string source = "/* banner */ @import /* before */ 'old.css' /* after */ layer(book) supports(display:grid) screen;\n" +
            "#heading/**/.lead { color:navy }";
        string result = HtmlResourcePipeline.RewriteCssResourceUrls(source, (url, _) => url == "old.css" ? "new.css" : url);
        Assert.Equal(source.Replace("'old.css'", "url(\"new.css\")"), result);
    }

    [Theory]
    [InlineData("url(images/*literal*/old.png)", "images/*literal*/old.png")]
    [InlineData("url('images/*literal*/old.png')", "images/*literal*/old.png")]
    [InlineData("u\\72 l(images/*literal*/old.png)", "images/*literal*/old.png")]
    public void CommentLikeTextInsideUrlTokensIsPartOfTheResource(string carrier, string expected) {
        string css = "/* url(fake.png) */ p{background:" + carrier + "}";
        var visited = new List<string>();
        string result = HtmlResourcePipeline.RewriteCssResourceUrls(css, (url, _) => { visited.Add(url); return "new.png"; });
        Assert.Equal(new[] { expected }, visited);
        Assert.Equal("/* url(fake.png) */ p{background:url(\"new.png\")}", result);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ChapterMergeAndAssetMovePreserveSelectorMeaningThroughReopen(bool merge) {
        var book = EpubPublication.Create("CSS preservation", "en");
        book.AddResource("image", "EPUB/old.svg", "image/svg+xml", Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='1' height='1'/>"));
        book.AddStylesheet("style", "EPUB/style.css", "#heading/**/.lead {color:#13579b;background:url('old.svg')} /* retain */");
        book.AddChapter("one", "EPUB/one.xhtml", "One", "<h1 id='heading' class='lead'>First</h1>", new[] { "style" });
        book.AddChapter("two", "EPUB/parts/two.xhtml", "Two", "<h1 id='second' class='lead'>Second</h1>", new[] { "style" });
        var second = book.GetContentXml("two"); XNamespace html = "http://www.w3.org/1999/xhtml";
        second.Root!.Element(html + "head")!.Add(new XElement(html + "style", "#second/**/.lead{color:#2468ac;background:url('../old.svg')}"));
        book.SetContentXml("two", second);
        if (merge) book.MergeChapters("one", "two", "second-start", new EpubChapterMergeOptions { StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles });
        else book.RenameResource("image", "EPUB/new.svg");
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal(merge ? "#heading/**/.lead {color:#13579b;background:url('old.svg')} /* retain */" :
            "#heading/**/.lead {color:#13579b;background:url(\"new.svg\")} /* retain */", Encoding.UTF8.GetString(reopened.GetResourceBytes("style")));
        string inline = reopened.GetContentXml(merge ? "one" : "two").Descendants(html + "style").Single().Value;
        Assert.Equal(merge ? "#second/**/.lead{color:#2468ac;background:url(\"old.svg\")}" :
            "#second/**/.lead{color:#2468ac;background:url(\"../new.svg\")}", inline);
    }
}
