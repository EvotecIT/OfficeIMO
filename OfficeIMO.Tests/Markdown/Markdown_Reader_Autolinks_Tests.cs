using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_Autolinks_Tests {
    [Fact]
    public void Autolinks_Http_Inside_Text() {
        var doc = MarkdownReader.Parse("See https://example.com.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"https://example.com\">https://example.com</a>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("https://example.com.</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Trim_Trailing_Bang_And_Question() {
        var doc = MarkdownReader.Parse("See https://example.com! And https://example.com?");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"https://example.com\">https://example.com</a>!", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"https://example.com\">https://example.com</a>?", html, StringComparison.Ordinal);
        Assert.DoesNotContain("https://example.com!</a>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("https://example.com?</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Keep_Balanced_Parentheses_In_Http_Urls() {
        var doc = MarkdownReader.Parse("See https://en.wikipedia.org/wiki/Function_(mathematics) and continue.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains(
            "<a href=\"https://en.wikipedia.org/wiki/Function_(mathematics)\">https://en.wikipedia.org/wiki/Function_(mathematics)</a>",
            html,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Http_Urls_After_Open_Paren() {
        var doc = MarkdownReader.Parse("See (https://en.wikipedia.org/wiki/Function_(mathematics)) now.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://en.wikipedia.org/wiki/Function_(mathematics)\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>See (https://en.wikipedia.org/wiki/Function_(mathematics)) now.</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Ambiguous_Paren_Suffixed_Urls() {
        var doc = MarkdownReader.Parse("Visit https://example.com/path_(x)).");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com/path_(x)\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit https://example.com/path_(x)).</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Balanced_Paren_Urls_With_Trailing_Comma() {
        var doc = MarkdownReader.Parse("Visit https://example.com/path_(x), ok");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com/path_(x)\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit https://example.com/path_(x), ok</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Www_Balanced_Paren_Urls_With_Trailing_Dot() {
        var doc = MarkdownReader.Parse("Visit www.example.com/path_(x).");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://www.example.com/path_(x)\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit www.example.com/path_(x).</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Www_Inside_Text() {
        var doc = MarkdownReader.Parse("See www.example.com, thanks.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"https://www.example.com\">www.example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Keep_Balanced_Parentheses_In_Www_Urls() {
        var doc = MarkdownReader.Parse("See www.example.com/path_(demo) next.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains(
            "<a href=\"https://www.example.com/path_(demo)\">www.example.com/path_(demo)</a>",
            html,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Http_Urls_With_Query_Parentheses() {
        var doc = MarkdownReader.Parse("Visit https://example.com/search?q=(x) now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com/search?q=(x)\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit https://example.com/search?q=(x) now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Www_Urls_With_Query_Parentheses() {
        var doc = MarkdownReader.Parse("Visit www.example.com/search?q=(x) now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://www.example.com/search?q=(x)\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit www.example.com/search?q=(x) now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Still_Link_Path_Parentheses_Before_Query_String() {
        var doc = MarkdownReader.Parse("Visit https://example.com/path_(demo)?q=value now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains(
            "<a href=\"https://example.com/path_(demo)?q=value\">https://example.com/path_(demo)?q=value</a>",
            html,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Http_Urls_With_Query_Ampersands() {
        var doc = MarkdownReader.Parse("Visit https://example.com/path?q=1&next=2 now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com/path?q=1&amp;next=2\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit https://example.com/path?q=1&amp;next=2 now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Gfm_Autolinks_Link_Query_And_Fragment_Special_Characters_With_Source_Metadata() {
        const string markdown = "Visit https://example.com/path?q=1&next=2 now\n";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var html = result.Document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        var paragraph = Assert.Single(result.SyntaxTree.Children);
        var link = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.InlineLink);
        var target = Assert.Single(link.Children, node => node.Kind == MarkdownSyntaxKind.InlineLinkTarget);
        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var nativeLink = Assert.Single(native.EnumerateInlines(), inline => inline.Kind == MarkdownNativeInlineKind.Link);
        var nativeTarget = Assert.Single(nativeLink.Metadata, metadata => metadata.Name == "target");

        Assert.Contains("<a href=\"https://example.com/path?q=1&amp;next=2\">https://example.com/path?q=1&amp;next=2</a>", html, StringComparison.Ordinal);
        Assert.Equal("https://example.com/path?q=1&next=2", link.Literal);
        Assert.Equal("https://example.com/path?q=1&next=2", target.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 41), target.SourceSpan);
        Assert.Equal("https://example.com/path?q=1&next=2", nativeTarget.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 41), nativeTarget.SourceSpan);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void Gfm_Autolinks_Render_Unicode_Http_Domain_As_Idn_While_Preserving_Source_Metadata() {
        const string markdown = "Visit https://пример.рф/path now\n";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var html = result.Document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        var paragraph = Assert.Single(result.SyntaxTree.Children);
        var link = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.InlineLink);
        var target = Assert.Single(link.Children, node => node.Kind == MarkdownSyntaxKind.InlineLinkTarget);
        var semanticParagraph = Assert.Single(result.Document.Blocks.OfType<ParagraphBlock>());
        var semanticLink = Assert.Single(semanticParagraph.Inlines.Nodes.OfType<LinkInline>());
        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var nativeLink = Assert.Single(native.EnumerateInlines(), inline => inline.Kind == MarkdownNativeInlineKind.Link);
        var nativeTarget = Assert.Single(nativeLink.Metadata, metadata => metadata.Name == "target");
        var written = result.Document.ToMarkdown().Replace("\r\n", "\n").Trim();

        Assert.Contains("<a href=\"https://xn--e1afmkfd.xn--p1ai/path\">https://пример.рф/path</a>", html, StringComparison.Ordinal);
        Assert.Equal("https://пример.рф/path", semanticLink.Text);
        Assert.Equal("https://пример.рф/path", semanticLink.Url);
        Assert.Equal("https://пример.рф/path", link.Literal);
        Assert.Equal("https://пример.рф/path", target.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 28), target.SourceSpan);
        Assert.Equal("https://пример.рф/path", nativeTarget.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 28), nativeTarget.SourceSpan);
        Assert.Equal("Visit https://пример.рф/path now", written);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void Gfm_Autolinks_Render_Unicode_Ftp_Domain_As_Idn_While_Preserving_Source_Metadata() {
        const string markdown = "Visit ftp://пример.рф/path now\n";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var html = result.Document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        var paragraph = Assert.Single(result.SyntaxTree.Children);
        var link = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.InlineLink);
        var target = Assert.Single(link.Children, node => node.Kind == MarkdownSyntaxKind.InlineLinkTarget);
        var semanticParagraph = Assert.Single(result.Document.Blocks.OfType<ParagraphBlock>());
        var semanticLink = Assert.Single(semanticParagraph.Inlines.Nodes.OfType<LinkInline>());
        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var nativeLink = Assert.Single(native.EnumerateInlines(), inline => inline.Kind == MarkdownNativeInlineKind.Link);
        var nativeTarget = Assert.Single(nativeLink.Metadata, metadata => metadata.Name == "target");
        var written = result.Document.ToMarkdown().Replace("\r\n", "\n").Trim();

        Assert.Contains("<a href=\"ftp://xn--e1afmkfd.xn--p1ai/path\">ftp://пример.рф/path</a>", html, StringComparison.Ordinal);
        Assert.Equal("ftp://пример.рф/path", semanticLink.Text);
        Assert.Equal("ftp://пример.рф/path", semanticLink.Url);
        Assert.Equal("ftp://пример.рф/path", link.Literal);
        Assert.Equal("ftp://пример.рф/path", target.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 26), target.SourceSpan);
        Assert.Equal("ftp://пример.рф/path", nativeTarget.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 26), nativeTarget.SourceSpan);
        Assert.Equal("Visit ftp://пример.рф/path now", written);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void Gfm_Autolinks_Render_Unicode_Http_Path_As_PercentEncoded_Href_While_Preserving_Display_And_Source() {
        const string markdown = "Visit https://example.com/ścieżka?q=zażółć now\n";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var html = result.Document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        var paragraph = Assert.Single(result.SyntaxTree.Children);
        var link = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.InlineLink);
        var target = Assert.Single(link.Children, node => node.Kind == MarkdownSyntaxKind.InlineLinkTarget);
        var semanticParagraph = Assert.Single(result.Document.Blocks.OfType<ParagraphBlock>());
        var semanticLink = Assert.Single(semanticParagraph.Inlines.Nodes.OfType<LinkInline>());
        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var nativeLink = Assert.Single(native.EnumerateInlines(), inline => inline.Kind == MarkdownNativeInlineKind.Link);
        var nativeTarget = Assert.Single(nativeLink.Metadata, metadata => metadata.Name == "target");
        var written = result.Document.ToMarkdown().Replace("\r\n", "\n").Trim();

        Assert.Contains("<a href=\"https://example.com/%C5%9Bcie%C5%BCka?q=za%C5%BC%C3%B3%C5%82%C4%87\">https://example.com/ścieżka?q=zażółć</a>", html, StringComparison.Ordinal);
        Assert.Equal("https://example.com/ścieżka?q=zażółć", semanticLink.Text);
        Assert.Equal("https://example.com/ścieżka?q=zażółć", semanticLink.Url);
        Assert.Equal("https://example.com/ścieżka?q=zażółć", link.Literal);
        Assert.Equal("https://example.com/ścieżka?q=zażółć", target.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 42), target.SourceSpan);
        Assert.Equal("https://example.com/ścieżka?q=zażółć", nativeTarget.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 42), nativeTarget.SourceSpan);
        Assert.Equal("Visit https://example.com/ścieżka?q=zażółć now", written);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void Gfm_Autolinks_Render_Unicode_Www_Domain_As_Idn_While_Preserving_Source_Literal() {
        const string markdown = "Visit www.пример.рф/path now\n";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var html = result.Document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        var paragraph = Assert.Single(result.SyntaxTree.Children);
        var link = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.InlineLink);
        var target = Assert.Single(link.Children, node => node.Kind == MarkdownSyntaxKind.InlineLinkTarget);
        var semanticParagraph = Assert.Single(result.Document.Blocks.OfType<ParagraphBlock>());
        var semanticLink = Assert.Single(semanticParagraph.Inlines.Nodes.OfType<LinkInline>());
        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var nativeLink = Assert.Single(native.EnumerateInlines(), inline => inline.Kind == MarkdownNativeInlineKind.Link);
        var nativeTarget = Assert.Single(nativeLink.Metadata, metadata => metadata.Name == "target");
        var written = result.Document.ToMarkdown().Replace("\r\n", "\n").Trim();

        Assert.Contains("<a href=\"http://www.xn--e1afmkfd.xn--p1ai/path\">www.пример.рф/path</a>", html, StringComparison.Ordinal);
        Assert.Equal("www.пример.рф/path", semanticLink.Text);
        Assert.Equal("http://www.пример.рф/path", semanticLink.Url);
        Assert.Equal("http://www.пример.рф/path", link.Literal);
        Assert.Equal("http://www.пример.рф/path", target.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 24), target.SourceSpan);
        Assert.Equal("http://www.пример.рф/path", nativeTarget.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 24), nativeTarget.SourceSpan);
        Assert.Equal("Visit www.пример.рф/path now", written);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void Gfm_Autolinks_Link_Balanced_Parentheses_Before_Trailing_Punctuation_With_Source_Metadata() {
        const string markdown = "Visit https://example.com/path_(x)). now\n";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var html = result.Document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        var paragraph = Assert.Single(result.SyntaxTree.Children);
        var link = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.InlineLink);
        var target = Assert.Single(link.Children, node => node.Kind == MarkdownSyntaxKind.InlineLinkTarget);
        var semanticParagraph = Assert.Single(result.Document.Blocks.OfType<ParagraphBlock>());
        var semanticLink = Assert.Single(semanticParagraph.Inlines.Nodes.OfType<LinkInline>());
        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var nativeLink = Assert.Single(native.EnumerateInlines(), inline => inline.Kind == MarkdownNativeInlineKind.Link);
        var nativeTarget = Assert.Single(nativeLink.Metadata, metadata => metadata.Name == "target");
        var written = result.Document.ToMarkdown().Replace("\r\n", "\n").Trim();

        Assert.Contains("<a href=\"https://example.com/path_(x)\">https://example.com/path_(x)</a>). now", html, StringComparison.Ordinal);
        Assert.Equal("https://example.com/path_(x)", semanticLink.Text);
        Assert.Equal("https://example.com/path_(x)", semanticLink.Url);
        Assert.Equal("https://example.com/path_(x)", link.Literal);
        Assert.Equal("https://example.com/path_(x)", target.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 34), target.SourceSpan);
        Assert.Equal("https://example.com/path_(x)", nativeTarget.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 34), nativeTarget.SourceSpan);
        Assert.Equal("Visit https://example.com/path_(x)\\). now", written);
        Assert.DoesNotContain("[https://example.com/path_(x)](https://example.com/path_(x))", written, StringComparison.Ordinal);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void Markdig_Autolinks_Keep_Trailing_Period_Before_Closing_Parenthesis_With_Source_Metadata() {
        const string markdown = "Visit https://example.com/path.) now\n";

        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        options.AutolinkAllowTrailingPunctuationBeforeClosingParenthesis = true;

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        var html = result.Document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        var paragraph = Assert.Single(result.SyntaxTree.Children);
        var link = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.InlineLink);
        var target = Assert.Single(link.Children, node => node.Kind == MarkdownSyntaxKind.InlineLinkTarget);
        var semanticParagraph = Assert.Single(result.Document.Blocks.OfType<ParagraphBlock>());
        var semanticLink = Assert.Single(semanticParagraph.Inlines.Nodes.OfType<LinkInline>());
        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeLink = Assert.Single(native.EnumerateInlines(), inline => inline.Kind == MarkdownNativeInlineKind.Link);
        var nativeTarget = Assert.Single(nativeLink.Metadata, metadata => metadata.Name == "target");
        var written = result.Document.ToMarkdown().Replace("\r\n", "\n").Trim();

        Assert.Contains("<a href=\"https://example.com/path.\">https://example.com/path.</a>) now", html, StringComparison.Ordinal);
        Assert.Equal("https://example.com/path.", semanticLink.Text);
        Assert.Equal("https://example.com/path.", semanticLink.Url);
        Assert.Equal("https://example.com/path.", link.Literal);
        Assert.Equal("https://example.com/path.", target.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 31), target.SourceSpan);
        Assert.Equal("https://example.com/path.", nativeTarget.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 31), nativeTarget.SourceSpan);
        Assert.Equal("Visit https://example.com/path.\\) now", written);
        Assert.DoesNotContain("[https://example.com/path.](https://example.com/path.)", written, StringComparison.Ordinal);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void Markdig_Autolinks_Keep_Trailing_Comma_Before_Closing_Parenthesis_With_Source_Metadata() {
        const string markdown = "Visit https://example.com/path,) now\n";

        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        options.AutolinkAllowTrailingPunctuationBeforeClosingParenthesis = true;

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        var html = result.Document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        var paragraph = Assert.Single(result.SyntaxTree.Children);
        var link = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.InlineLink);
        var target = Assert.Single(link.Children, node => node.Kind == MarkdownSyntaxKind.InlineLinkTarget);
        var semanticParagraph = Assert.Single(result.Document.Blocks.OfType<ParagraphBlock>());
        var semanticLink = Assert.Single(semanticParagraph.Inlines.Nodes.OfType<LinkInline>());
        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeLink = Assert.Single(native.EnumerateInlines(), inline => inline.Kind == MarkdownNativeInlineKind.Link);
        var nativeTarget = Assert.Single(nativeLink.Metadata, metadata => metadata.Name == "target");
        var written = result.Document.ToMarkdown().Replace("\r\n", "\n").Trim();

        Assert.Contains("<a href=\"https://example.com/path,\">https://example.com/path,</a>) now", html, StringComparison.Ordinal);
        Assert.Equal("https://example.com/path,", semanticLink.Text);
        Assert.Equal("https://example.com/path,", semanticLink.Url);
        Assert.Equal("https://example.com/path,", link.Literal);
        Assert.Equal("https://example.com/path,", target.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 31), target.SourceSpan);
        Assert.Equal("https://example.com/path,", nativeTarget.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 31), nativeTarget.SourceSpan);
        Assert.Equal("Visit https://example.com/path,\\) now", written);
        Assert.DoesNotContain("[https://example.com/path,](https://example.com/path,)", written, StringComparison.Ordinal);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Theory]
    [InlineData("Visit https://example.com/path;) now", "<a href=\"https://example.com/path;\">https://example.com/path;</a>) now")]
    [InlineData("Visit https://example.com/path!) now", "<a href=\"https://example.com/path!\">https://example.com/path!</a>) now")]
    [InlineData("Visit https://example.com/path?) now", "<a href=\"https://example.com/path?\">https://example.com/path?</a>) now")]
    public void Markdig_Autolinks_Keep_Trailing_NonPeriod_Punctuation_Before_Closing_Parenthesis(string markdown, string expectedHtml) {
        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        options.AutolinkAllowTrailingPunctuationBeforeClosingParenthesis = true;

        var doc = MarkdownReader.Parse(markdown, options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains(expectedHtml, html, StringComparison.Ordinal);
    }

    [Fact]
    public void Gfm_Autolinks_Require_Lowercase_Www_Prefix_But_Allow_Mixed_Case_Host() {
        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        var upperPrefix = MarkdownReader.Parse("Visit WWW.example.com now", options)
            .ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        var mixedHost = MarkdownReader.Parse("Visit www.Example.com now", options)
            .ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=", upperPrefix, StringComparison.Ordinal);
        Assert.Contains("<p>Visit WWW.example.com now</p>", upperPrefix, StringComparison.Ordinal);
        Assert.Contains("<a href=\"http://www.Example.com\">www.Example.com</a>", mixedHost, StringComparison.Ordinal);
    }

    [Fact]
    public void Gfm_Autolinks_Require_Lowercase_Bare_Scheme_Prefixes() {
        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        var html = MarkdownReader
            .Parse("Fetch FTP://example.com/file and ftp://example.com/file; call TEL:+123 and tel:+123; mail MAILTO:user@example.com and mailto:user@example.com.", options)
            .ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"FTP://example.com/file\"", html, StringComparison.Ordinal);
        Assert.Contains("FTP://example.com/file", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"ftp://example.com/file\">ftp://example.com/file</a>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("href=\"TEL:+123\"", html, StringComparison.Ordinal);
        Assert.Contains("TEL:+123", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"tel:+123\">+123</a>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("href=\"MAILTO:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("MAILTO:user@example.com", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"mailto:user@example.com\">mailto:user@example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Gfm_Autolinks_Can_Narrow_Bare_Scheme_Prefixes_For_Markdig_Compatibility() {
        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        options.AutolinkBareSchemePrefixes = new[] { "mailto:", "ftp://", "tel:" };

        var html = MarkdownReader
            .Parse("Use mailto:user@example.com, ftp://example.com/file, tel:+123, and xmpp:user@example.com.", options)
            .ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"mailto:user@example.com\">mailto:user@example.com</a>", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"ftp://example.com/file\">ftp://example.com/file</a>", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"tel:+123\">+123</a>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("href=\"xmpp:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("xmpp:user@example.com", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdig_Autolinks_Can_Display_Bare_Mailto_As_Address_Only_While_Writing_Source_Literal() {
        const string markdown = "Contact mailto:user@example.com now\n";
        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        options.AutolinkBareMailtoDisplayAddressOnly = true;

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        var html = result.Document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        var paragraph = Assert.Single(result.SyntaxTree.Children);
        var link = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.InlineLink);
        var target = Assert.Single(link.Children, node => node.Kind == MarkdownSyntaxKind.InlineLinkTarget);
        var semanticParagraph = Assert.Single(result.Document.Blocks.OfType<ParagraphBlock>());
        var semanticLink = Assert.Single(semanticParagraph.Inlines.Nodes.OfType<LinkInline>());
        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeLink = Assert.Single(native.EnumerateInlines(), inline => inline.Kind == MarkdownNativeInlineKind.Link);
        var nativeTarget = Assert.Single(nativeLink.Metadata, metadata => metadata.Name == "target");
        var written = result.Document.ToMarkdown().Replace("\r\n", "\n").Trim();

        Assert.Contains("<a href=\"mailto:user@example.com\">user@example.com</a>", html, StringComparison.Ordinal);
        Assert.Equal("user@example.com", semanticLink.Text);
        Assert.Equal("mailto:user@example.com", semanticLink.Url);
        Assert.Equal("mailto:user@example.com", link.Literal);
        Assert.Equal("mailto:user@example.com", target.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 9, 1, 31), target.SourceSpan);
        Assert.Equal("mailto:user@example.com", nativeTarget.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 9, 1, 31), nativeTarget.SourceSpan);
        Assert.Equal("Contact mailto:user@example.com now", written);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Http_Urls_With_Fragment_Ampersands() {
        var doc = MarkdownReader.Parse("Visit https://example.com/path#frag&next now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com/path#frag&amp;next\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit https://example.com/path#frag&amp;next now</p>", html, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Visit _https://example.com now")]
    [InlineData("Visit /https://example.com now")]
    [InlineData("Visit &https://example.com now")]
    public void Autolinks_DoNot_Link_Http_Urls_After_Invalid_Left_Boundaries(string markdown) {
        var doc = MarkdownReader.Parse(markdown);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com\"", html, StringComparison.Ordinal);
        Assert.Contains(markdown.Replace("&", "&amp;"), html, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Visit foo:https://example.com now")]
    [InlineData("Visit foo.https://example.com now")]
    [InlineData("Visit foo+https://example.com now")]
    [InlineData("Visit foo-https://example.com now")]
    [InlineData("Visit foo=https://example.com now")]
    [InlineData("Visit (https://example.com now")]
    [InlineData("Visit (https://example.com) now")]
    [InlineData("Visit [https://example.com now")]
    public void Autolinks_DoNot_Link_Http_Urls_After_Common_Prefix_Punctuation(string markdown) {
        var doc = MarkdownReader.Parse(markdown);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com\"", html, StringComparison.Ordinal);
        Assert.Contains(markdown, html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Http_Urls_After_Apostrophe() {
        var doc = MarkdownReader.Parse("Visit 'https://example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit &#39;https://example.com now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Www_Urls_With_Query_Ampersands() {
        var doc = MarkdownReader.Parse("Visit www.example.com/path?q=1&next=2 now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://www.example.com/path?q=1&amp;next=2\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit www.example.com/path?q=1&amp;next=2 now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Www_Urls_After_Invalid_Left_Boundaries() {
        var doc = MarkdownReader.Parse("Visit _www.example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://www.example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit _www.example.com now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Www_Urls_After_Ampersand() {
        var doc = MarkdownReader.Parse("Visit &www.example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://www.example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit &amp;www.example.com now</p>", html, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Visit foo:www.example.com now")]
    [InlineData("Visit foo+www.example.com now")]
    [InlineData("Visit foo-www.example.com now")]
    [InlineData("Visit foo=www.example.com now")]
    [InlineData("Visit (www.example.com now")]
    [InlineData("Visit (www.example.com) now")]
    [InlineData("Visit [www.example.com now")]
    public void Autolinks_DoNot_Link_Www_Urls_After_Common_Prefix_Punctuation(string markdown) {
        var doc = MarkdownReader.Parse(markdown);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://www.example.com\"", html, StringComparison.Ordinal);
        Assert.Contains(markdown, html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Www_Urls_After_Apostrophe() {
        var doc = MarkdownReader.Parse("Visit 'www.example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://www.example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit &#39;www.example.com now</p>", html, StringComparison.Ordinal);
    }


    [Fact]
    public void Autolinks_Email_Inside_Text() {
        var doc = MarkdownReader.Parse("Email user@example.com.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"mailto:user@example.com\">user@example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Mailto_Email_Tokens() {
        var doc = MarkdownReader.Parse("Contact mailto:user@example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact mailto:user@example.com now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Emails_After_Slash() {
        var doc = MarkdownReader.Parse("Contact /user@example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact /user@example.com now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Emails_After_Colon() {
        var doc = MarkdownReader.Parse("Contact foo:user@example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact foo:user@example.com now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Emails_After_Equals() {
        var doc = MarkdownReader.Parse("Contact foo=user@example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact foo=user@example.com now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Emails_After_Open_Bracket() {
        var doc = MarkdownReader.Parse("Contact [user@example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact [user@example.com now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Emails_After_Ampersand() {
        var doc = MarkdownReader.Parse("Contact &user@example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact &amp;user@example.com now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Emails_After_Open_Paren() {
        var doc = MarkdownReader.Parse("Contact (user@example.com) now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact (user@example.com) now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Gfm_Autolinks_DoNot_Crash_On_Upstream_Ignored_Email_Case() {
        const string markdown = "This shouldn't crash everything: (_A_@_.A";

        var result = MarkdownReader.ParseWithSyntaxTree(
            markdown,
            MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var html = result.Document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("This shouldn&#39;t crash everything", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<a ", html, StringComparison.OrdinalIgnoreCase);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Emails_After_Apostrophe() {
        var doc = MarkdownReader.Parse("Contact 'user@example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact &#39;user@example.com now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Emails_With_Path_Suffixes() {
        var doc = MarkdownReader.Parse("Contact user@example.com/path now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact user@example.com/path now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Emails_With_Fragment_Suffixes() {
        var doc = MarkdownReader.Parse("Contact user@example.com#frag now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact user@example.com#frag now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Link_Plain_Emails_With_Plus_Tags() {
        var doc = MarkdownReader.Parse("Contact user.name+tag@example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains(
            "<a href=\"mailto:user.name+tag@example.com\">user.name+tag@example.com</a>",
            html,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Angle_Autolinks_Explicit_Mailto_Are_Supported() {
        var doc = MarkdownReader.Parse("Contact <mailto:user@example.com>.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"mailto:user@example.com\">mailto:user@example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Angle_Autolinks_Explicit_Mailto_Respect_Url_Policy() {
        var options = new MarkdownReaderOptions {
            RestrictUrlSchemes = true,
            AllowedUrlSchemes = new[] { "http", "https" }
        };
        var doc = MarkdownReader.Parse("Contact <mailto:user@example.com>.", options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("&lt;mailto:user@example.com&gt;", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Angle_Autolinks_Explicit_Absolute_Uris_Are_Supported() {
        var doc = MarkdownReader.Parse("Fetch <ftp://example.com/file.txt>.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"ftp://example.com/file.txt\">ftp://example.com/file.txt</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Angle_Autolinks_Explicit_Tel_Uris_Are_Supported() {
        var doc = MarkdownReader.Parse("Call <tel:+123456789>.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"tel:+123456789\">tel:+123456789</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Angle_Autolinks_Explicit_Urn_Uris_Are_Supported() {
        var doc = MarkdownReader.Parse("Lookup <urn:isbn:9780143127741>.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"urn:isbn:9780143127741\">urn:isbn:9780143127741</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Angle_Autolinks_Explicit_Absolute_Uris_Respect_Url_Policy() {
        var options = new MarkdownReaderOptions {
            RestrictUrlSchemes = true,
            AllowedUrlSchemes = new[] { "http", "https" }
        };
        var doc = MarkdownReader.Parse("Fetch <ftp://example.com/file.txt>.", options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"ftp://example.com/file.txt\"", html, StringComparison.Ordinal);
        Assert.Contains("&lt;ftp://example.com/file.txt&gt;", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Angle_Autolinks_Explicit_NonHierarchical_Uris_Respect_Url_Policy() {
        var options = new MarkdownReaderOptions {
            RestrictUrlSchemes = true,
            AllowedUrlSchemes = new[] { "http", "https" }
        };
        var doc = MarkdownReader.Parse("Call <tel:+123456789>.", options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"tel:+123456789\"", html, StringComparison.Ordinal);
        Assert.Contains("&lt;tel:+123456789&gt;", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Bare_Scheme_Autolinks_Are_Opt_In_For_Ftp_And_Tel() {
        var doc = MarkdownReader.Parse("Fetch ftp://example.com/file.txt and call tel:+123456789.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"ftp://example.com/file.txt\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("href=\"tel:+123456789\"", html, StringComparison.Ordinal);
        Assert.Contains("Fetch ftp://example.com/file.txt and call tel:+123456789.", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Gfm_Autolinks_Link_Markdig_Ftp_And_Tel_Bare_Schemes() {
        const string markdown = "Fetch ftp://example.com/file.txt and call tel:+123-456.";

        var doc = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"ftp://example.com/file.txt\">ftp://example.com/file.txt</a>", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"tel:+123-456\">+123-456</a>.", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Gfm_Autolinks_Reject_Ftp_Domain_Without_Period() {
        var doc = MarkdownReader.Parse("Fetch ftp://localhost/file and ftp://example.com/file", MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"ftp://localhost/file\"", html, StringComparison.Ordinal);
        Assert.Contains("ftp://localhost/file", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"ftp://example.com/file\">ftp://example.com/file</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Work_In_Tables_And_Lists() {
        var md = """
| Link |
| --- |
| www.example.com |

- Email user@example.com
""";
        var doc = MarkdownReader.Parse(md);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"https://www.example.com\">www.example.com</a>", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"mailto:user@example.com\">user@example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Can_Require_Domain_Period_For_Markdig_Style_Compatibility() {
        var options = new MarkdownReaderOptions {
            AutolinkAllowDomainWithoutPeriod = false
        };

        var doc = MarkdownReader.Parse(
            "See https://localhost and www.local and https://example.com and www.example.com",
            options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://localhost\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("href=\"https://www.local\"", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"https://example.com\">https://example.com</a>", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"https://www.example.com\">www.example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Default_Profile_Preserves_Domain_Without_Period_Legacy_Behavior() {
        var doc = MarkdownReader.Parse("See https://localhost and www.local");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"https://localhost\">https://localhost</a>", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"https://www.local\">www.local</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_ValidPreviousCharacters_Can_Use_Markdig_Style_Boundaries() {
        var options = new MarkdownReaderOptions {
            AutolinkValidPreviousCharacters = "_('"
        };

        var doc = MarkdownReader.Parse("See _https://example.com and (www.example.com) and 'user@example.com", options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("_<a href=\"https://example.com\">https://example.com</a>", html, StringComparison.Ordinal);
        Assert.Contains("(<a href=\"https://www.example.com\">www.example.com</a>)", html, StringComparison.Ordinal);
        Assert.Contains("&#39;<a href=\"mailto:user@example.com\">user@example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Table_Cells_Respect_Domain_Period_Option() {
        const string markdown = """
| Link |
| --- |
| https://localhost |
| https://example.com |
""";
        var options = new MarkdownReaderOptions {
            AutolinkAllowDomainWithoutPeriod = false
        };

        var doc = MarkdownReader.Parse(markdown, options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://localhost\"", html, StringComparison.Ordinal);
        Assert.Contains("<td>https://localhost</td>", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"https://example.com\">https://example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Can_Be_Disabled() {
        var options = new MarkdownReaderOptions {
            AutolinkUrls = false,
            AutolinkWwwUrls = false,
            AutolinkEmails = false
        };
        var doc = MarkdownReader.Parse("See https://example.com and www.example.com and user@example.com", options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("href=\"https://www.example.com\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Can_Use_Portable_Profile() {
        var options = MarkdownReaderOptions.CreatePortableProfile();
        var doc = MarkdownReader.Parse("See https://example.com and www.example.com and user@example.com and <angle@example.com>", options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("href=\"https://www.example.com\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("href=\"mailto:user@example.com\">user@example.com</a>", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"mailto:angle@example.com\">angle@example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Portable_Profile_Disables_Callouts_And_Task_Checkboxes() {
        var options = MarkdownReaderOptions.CreatePortableProfile();

        var callout = MarkdownReader.Parse("> [!NOTE]\n> body", options)
            .ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        var taskList = MarkdownReader.Parse("- [ ] task", options)
            .ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("class=\"callout", callout, StringComparison.Ordinal);
        Assert.Contains("[!NOTE]", callout, StringComparison.Ordinal);
        Assert.DoesNotContain("task-list-item-checkbox", taskList, StringComparison.Ordinal);
        Assert.Contains("[ ] task", taskList, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Require_Left_Boundary() {
        var doc = MarkdownReader.Parse("prefixhttps://example.com should not linkify.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_RenderMarkdown_Preserves_Source_Backed_Bare_And_Angle_Spelling() {
        const string markdown = "See https://example.com and www.example.com and user@example.com and <https://example.com/docs> and mailto:team@example.com and ftp://example.com/file.txt and tel:+123-456.";

        var doc = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var written = doc.ToMarkdown().Replace("\r\n", "\n").Trim();

        Assert.Equal(markdown, written);
        Assert.DoesNotContain("[https://example.com](https://example.com)", written, StringComparison.Ordinal);
        Assert.DoesNotContain("[www.example.com](http://www.example.com)", written, StringComparison.Ordinal);
        Assert.DoesNotContain("[user@example.com](mailto:user@example.com)", written, StringComparison.Ordinal);
        Assert.Contains("<https://example.com/docs>", written, StringComparison.Ordinal);
        Assert.Contains("mailto:team@example.com", written, StringComparison.Ordinal);
        Assert.Contains("ftp://example.com/file.txt", written, StringComparison.Ordinal);
        Assert.Contains("tel:+123-456", written, StringComparison.Ordinal);
    }

    [Fact]
    public void Explicit_Links_RenderMarkdown_Do_Not_Become_Autolinks() {
        const string markdown = "[https://example.com](https://example.com) and [www.example.com](https://www.example.com)";

        var doc = MarkdownReader.Parse(markdown);
        var written = doc.ToMarkdown().Replace("\r\n", "\n").Trim();

        Assert.Equal(markdown, written);
    }
}
