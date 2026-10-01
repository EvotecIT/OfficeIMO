namespace OfficeIMO.Email.Tests;

public sealed class EmailIndexTextTests {
    [Fact]
    public void SharedBodySourceBudgetRejectsBeforeGeneratedHtmlIsParsed() {
        var source = new EmailDocument();
        source.Body.Text = new string('&', 100);
        var policy = OfficeIMO.Html.HtmlConversionDocumentOptions.CreateUntrustedProfile();
        policy.Limits.MaxInputCharacters = 1;
        Assert.Throws<ArgumentException>(() => EmailBodyProjection.Create(source, new EmailBodyProjectionOptions {
            MaxBodySourceCharacters = 3, HtmlOptions = policy
        }));
    }

    [Fact]
    public void PlainTextRegionsRetainQuotesAndSignatureUntilExplicitlyExcluded() {
        var source = new EmailDocument();
        source.Body.Text = "Hello\r\n> Previous reply\r\n-- \r\nAlex";
        var full = EmailIndexText.Create(source);
        Assert.Equal("Hello\n> Previous reply\n-- \nAlex\n", full.FullText);
        Assert.Equal(full.FullText, full.SelectedText);
        Assert.Contains(full.Regions, region => region.Kind == EmailIndexRegionKind.Quoted && region.Reason == "plain-quote-prefix");
        Assert.Contains(full.Regions, region => region.Kind == EmailIndexRegionKind.Signature && region.Reason == "plain-signature-separator");
        var selected = EmailIndexText.Create(source, new EmailIndexTextOptions { ExcludeQuotes = true, ExcludeSignatures = true });
        Assert.Equal(full.FullText, selected.FullText);
        Assert.Equal("Hello\n", selected.SelectedText);
        Assert.Equal("Hello\r\n> Previous reply\r\n-- \r\nAlex", source.Body.Text);
        VerifyCoverage(selected);
    }

    [Fact]
    public void HtmlUsesOwnedSafeProjectionAndMaintainsInlineAndBlockBoundaries() {
        var source = new EmailDocument();
        source.Body.Html = "<p>Hello <b>world</b> &amp; friends</p><div class='gmail_quote'>Old <b>reply</b></div>" +
            "<div class='gmail_signature'>Alex</div><script>secret-script</script><style>secret-style</style>";
        var result = EmailIndexText.Create(source, new EmailIndexTextOptions { ExcludeQuotes = true, ExcludeSignatures = true });
        Assert.Equal("Hello world & friends\nOld reply\nAlex\n", result.FullText);
        Assert.Equal("Hello world & friends\n", result.SelectedText);
        Assert.Contains(result.Regions, region => region.Reason == "gmail_quote");
        Assert.Contains(result.Regions, region => region.Reason == "gmail_signature");
        Assert.Equal(EmailBodySourceKind.Html, result.SourceKind);
        VerifyCoverage(result);
    }

    [Theory]
    [InlineData(4, "abc")]
    [InlineData(5, "abc\ud83d\ude00")]
    public void ClippingPreservesScalarBoundariesAndRegionOffsets(int maximum, string expected) {
        var source = new EmailDocument();
        source.Body.Text = "abc\ud83d\ude00def";
        var result = EmailIndexText.Create(source, new EmailIndexTextOptions { MaxTextChars = maximum });
        Assert.True(result.Truncated);
        Assert.Equal(expected, result.FullText);
        Assert.Contains(result.Diagnostics, value => value.Code == "EMAIL_INDEX_TEXT_TRUNCATED");
        VerifyCoverage(result);
    }

    [Fact]
    public void PrefersPlainTextAndRejectsOversizedSelectedInputBeforeParsing() {
        var source = new EmailDocument();
        source.Body.Text = "plain";
        source.Body.Html = "<p>HTML alternative</p>";
        Assert.Equal("plain\n", EmailIndexText.Create(source).FullText);
        Assert.Equal("HTML alternative\n", EmailIndexText.Create(source, new EmailIndexTextOptions { PreferPlainText = false }).FullText);
        Assert.Throws<ArgumentException>(() => EmailIndexText.Create(source, new EmailIndexTextOptions { MaxSourceChars = 4 }));
        Assert.Throws<ArgumentException>(() => EmailIndexText.Create(source, new EmailIndexTextOptions { PreferPlainText = false, MaxSourceChars = 8 }));
    }

    [Fact]
    public void BodyOnlyProjectionDoesNotOpenInlineContentAndQuoteAncestryWins() {
        var source = new EmailDocument();
        source.Body.Html = "<blockquote>Old <span class='moz-signature'>Signature</span></blockquote><p>New</p><img src='cid:logo'>";
        source.Attachments.Add(new EmailAttachment { ContentId = "logo", IsInline = true, ContentSource = new ForbiddenSource() });
        var result = EmailIndexText.Create(source, new EmailIndexTextOptions { ExcludeQuotes = true });
        Assert.Equal("New\n", result.SelectedText);
        Assert.DoesNotContain(result.Regions, value => value.Kind == EmailIndexRegionKind.Signature);
        VerifyCoverage(result);
    }

    [Fact]
    public void ExcludedInlineSignatureDoesNotJoinAdjacentIndexWords() {
        var source = new EmailDocument();
        source.Body.Html = "<p>Before<span class='gmail_signature'>Private signature</span>After</p>";
        var result = EmailIndexText.Create(source, new EmailIndexTextOptions { ExcludeSignatures = true });
        Assert.Equal("Before\nAfter\n", result.SelectedText);
        Assert.Equal("BeforePrivate signatureAfter\n", result.FullText);
        VerifyCoverage(result);
    }

    private static void VerifyCoverage(EmailIndexTextResult result) {
        int offset = 0;
        foreach (var region in result.Regions) {
            Assert.Equal(offset, region.Start);
            Assert.True(region.Length > 0);
            offset += region.Length;
        }
        Assert.Equal(result.FullText.Length, offset);
    }
    private sealed class ForbiddenSource : IEmailContentSource {
        public long? Length => null;
        public Stream OpenRead() => throw new InvalidOperationException("Indexing must not open attachments.");
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) => throw new InvalidOperationException("Indexing must not open attachments.");
    }
}
