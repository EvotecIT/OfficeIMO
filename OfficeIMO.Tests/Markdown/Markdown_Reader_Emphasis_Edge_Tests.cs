using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_Emphasis_Edge_Tests {
    [Fact]
    public void Triple_Closer_Can_Close_Italic_Then_Bold() {
        var md = "**bold *italic***";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<strong>bold <em>italic</em></strong>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Triple_Marker_Run_Can_Render_Italic_With_Inner_Bold() {
        var md = "***foo***";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<em><strong>foo</strong></em>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<strong><em>foo</em></strong>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Triple_Marker_Run_Can_Keep_Italic_Open_After_Inner_Bold_Closes() {
        var md = "***foo** bar*";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<em><strong>foo</strong> bar</em>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("**<em>foo</em>* bar*", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Single_Star_Inside_Bold_Can_Remain_Literal_When_Only_Bold_Closer_Remains() {
        var md = "**foo*bar**";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<strong>foo*bar</strong>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("**foo<em>bar</em>*", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Double_Star_Inside_Italic_Can_Remain_Literal_When_Only_Italic_Closer_Remains() {
        var md = "*foo**bar*baz**";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<em>foo**bar</em>baz**", html, StringComparison.Ordinal);
        Assert.DoesNotContain("*foo<strong>bar*baz</strong>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Unclosed_Emphasis_Is_Literal() {
        var md = "*not closed";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("*not closed", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<em>", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Intraword_Underscores_Do_Not_Create_Emphasis() {
        var md = "foo_bar_baz";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("foo_bar_baz", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<em>", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Strikethrough_Can_Nest_Inside_Emphasis() {
        var md = "*a ~~b~~ c*";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<em>a <del>b</del> c</em>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Double_Star_Inside_Italic_Opens_Inner_Bold() {
        var md = "*foo**bar**baz*";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<em>foo<strong>bar</strong>baz</em>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<em>foo</em><em>bar</em><em>baz</em>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Double_Underscore_At_Start_Can_Degrade_To_Literal_Then_Italic() {
        var md = "__foo_bar_";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("_<em>foo_bar</em>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("__foo_bar_", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Double_Underscore_At_Start_Can_Degrade_When_Only_Single_Closer_Is_Valid() {
        var md = "__foo__bar_";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("_<em>foo__bar</em>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("__foo__bar_", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Double_Underscore_Still_Forms_Bold_When_A_Valid_Double_Closer_Exists() {
        var md = "__foo__ bar_";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<strong>foo</strong> bar_", html, StringComparison.Ordinal);
        Assert.DoesNotContain("_<em>foo__ bar</em>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Single_Underscore_Inside_Star_Italic_Can_Remain_Literal_When_Outer_Closer_Comes_First() {
        var md = "*a _b* c_";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<em>a _b</em> c_", html, StringComparison.Ordinal);
        Assert.DoesNotContain("*a <em>b* c</em>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Single_Star_Inside_Underscore_Italic_Can_Remain_Literal_When_Outer_Closer_Comes_First() {
        var md = "_a *b_ c*";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<em>a *b</em> c*", html, StringComparison.Ordinal);
        Assert.DoesNotContain("_a <em>b_ c</em>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Mixed_Markers_Can_Still_Nest_When_Inner_Closer_Comes_First() {
        var md = "*a _b_ c*";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<em>a <em>b</em> c</em>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Mixed_Markers_Can_Still_Nest_In_Reverse_Order_When_Inner_Closer_Comes_First() {
        var md = "_a *b* c_";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<em>a <em>b</em> c</em>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Double_Star_Inside_Italic_Can_Rebalance_Into_Dual_Italic_When_Single_Close_Comes_First() {
        var md = "*a **b* c**";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<em>a <em><em>b</em> c</em></em>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<em>a **b</em> c**", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Double_Underscore_Inside_Italic_Can_Rebalance_Into_Dual_Italic_When_Single_Close_Comes_First() {
        var md = "_a __b_ c__";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<em>a <em><em>b</em> c</em></em>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<em>a __b</em> c__", html, StringComparison.Ordinal);
    }
}

