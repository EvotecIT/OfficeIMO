namespace OfficeIMO.Email.Tests;

public sealed class EmailHtmlCompositionTests {
    [Fact]
    public void OutlookConditionalCommentsCannotCarryResourcesIntoAnIndependentDraft() {
        var source = new EmailDocument();
        source.Body.Html = "<p>Original<!--[if mso]><v:rect><v:fill src='https://example.test/pixel'/></v:rect><![endif]--></p>";
        var draft = EmailHtmlComposer.Reply(source, new EmailAddress("me@example.test"), "Reply").Document;
        Assert.DoesNotContain("example.test/pixel", draft.Body.Html);
        Assert.DoesNotContain("<!--[if", draft.Body.Html);
        using var bytes = new MemoryStream();
        Assert.False(new EmailDocumentWriter().Write(draft, bytes).HasErrors);
        bytes.Position = 0;
        using var read = new EmailDocumentReader().Read(bytes);
        Assert.DoesNotContain("example.test/pixel", read.Document.Body.Html);
        Assert.Contains("Original", read.Document.Body.Html);
    }

    [Fact]
    public void QuotationRemovesAutomaticRefreshInTheSharedMailSafetyProjection() {
        var source = new EmailDocument();
        source.Body.Html = "<p>Original</p><meta http-equiv='refresh' content='0; url=https://example.test/redirect'>";
        Assert.DoesNotContain("http-equiv", EmailBodyProjection.Create(source).Html);
        Assert.DoesNotContain("http-equiv", EmailHtmlComposer.Reply(source, new EmailAddress("me@example.test"), "Reply").Document.Body.Html);
    }

    [Theory]
    [InlineData("<input type='image' src='data:image/png;base64,iVBORw0KGgo='>")]
    [InlineData("<table background='data:image/png;base64,iVBORw0KGgo='><tr><td>Cell</td></tr></table>")]
    public void QuotationOmitsResourceAttributesAndImageInputsWithEvidence(string html) {
        var source = new EmailDocument();
        source.Body.Html = html;
        var result = EmailHtmlComposer.Reply(source, new EmailAddress("me@example.test"), "Reply");
        Assert.DoesNotContain("data:image", result.Document.Body.Html);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_COMPOSITION_RESOURCES_OMITTED");
    }

    [Theory]
    [InlineData("abc")]
    [InlineData("&&&")]
    public void PlainSourceCharacterBudgetDoesNotBoundItsEncodedIntermediate(string body) {
        var source = new EmailDocument();
        source.Body.Text = body;
        var result = EmailHtmlComposer.Reply(source, new EmailAddress("me@example.test"), "Reply", new EmailHtmlCompositionOptions { MaxSourceChars = 3 });
        Assert.Contains("> " + body, result.Document.Body.Text);
    }

    [Fact]
    public void BodylessMessagesDoNotAcquireFabricatedIndexOrQuoteText() {
        var source = new EmailDocument();
        var index = EmailIndexText.Create(source);
        Assert.Equal(EmailBodySourceKind.None, index.SourceKind);
        Assert.Empty(index.FullText);
        Assert.Empty(index.Regions);
        Assert.Contains(index.Diagnostics, value => value.Code == "EMAIL_BODY_MISSING");
        var result = EmailHtmlComposer.Reply(source, new EmailAddress("me@example.test"), "Reply");
        Assert.Equal("Reply", result.Document.Body.Text);
        Assert.DoesNotContain("<blockquote", result.Document.Body.Html);
    }

    [Fact]
    public void LargeEncodedPlainBodyReachesTextFallbackBeforeRichProjection() {
        var source = new EmailDocument();
        source.Body.Text = new string('&', 500_000);
        var result = EmailHtmlComposer.Reply(source, new EmailAddress("me@example.test"), "Reply");
        Assert.Contains(result.Diagnostics, value => value.Code == "EMAIL_COMPOSITION_TEXT_QUOTE_FALLBACK");
        Assert.Equal("Reply\r\n\r\n> " + new string('&', 256 * 1024), result.Document.Body.Text);
    }

    [Fact]
    public void RtfOriginalAndGeneratedHtmlHaveSeparateCharacterBudgets() {
        var source = new EmailDocument();
        source.Body.Rtf = "{\\rtf1 abc}";
        var index = EmailIndexText.Create(source, new EmailIndexTextOptions { MaxSourceChars = source.Body.Rtf.Length, PreferPlainText = false });
        Assert.Equal(EmailBodySourceKind.Rtf, index.SourceKind);
        Assert.Contains("abc", index.FullText);
        var result = EmailHtmlComposer.Reply(source, new EmailAddress("me@example.test"), "Reply", new EmailHtmlCompositionOptions { MaxSourceChars = source.Body.Rtf.Length });
        Assert.Contains("> abc", result.Document.Body.Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RichRepliesEncodeAuthoredTextOmitResourcesAndReuseRecipientThreadingPolicy(bool all) {
        var original = new EmailDocument { From = new EmailAddress("author@example.test"), Subject = "Topic", MessageId = "<parent@example.test>" };
        original.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("me@example.test")));
        original.Recipients.Add(new EmailRecipient(EmailRecipientKind.Cc, new EmailAddress("other@example.test")));
        original.Recipients.Add(new EmailRecipient(EmailRecipientKind.Bcc, new EmailAddress("private@example.test")));
        original.Headers.Add(new EmailHeader("DKIM-Signature", "original"));
        original.Body.Html = "<p style='color:red'>Original <b>rich</b> body</p><img src='cid:logo'><script>alert(1)</script>";
        var result = all ? EmailHtmlComposer.ReplyAll(original, new EmailAddress("me@example.test"), "<hello>\nNext")
            : EmailHtmlComposer.Reply(original, new EmailAddress("me@example.test"), "<hello>\nNext");
        string html = result.Document.Body.Html!;
        Assert.Contains("&lt;hello&gt;<br>Next", html);
        Assert.Contains("<b>rich</b>", html);
        Assert.DoesNotContain("<script", html);
        Assert.DoesNotContain("<img", html);
        Assert.DoesNotContain("style=", html);
        Assert.Equal("<parent@example.test>", result.Document.MessageMetadata.InReplyToId);
        Assert.Equal(all ? 2 : 1, result.Document.Recipients.Count);
        Assert.DoesNotContain(result.Document.Recipients, recipient => recipient.Kind == EmailRecipientKind.Bcc || recipient.Address.Address == "me@example.test");
        Assert.Contains("> Original rich body", result.Document.Body.Text);
        Assert.Empty(result.Document.Headers);
        Assert.Empty(result.Document.Attachments);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_COMPOSITION_RESOURCES_OMITTED");
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_COMPOSITION_PLAIN_BODY_UNAVAILABLE");
        using var bytes = new MemoryStream();
        var write = new EmailDocumentWriter().Write(result.Document, bytes);
        Assert.False(write.HasErrors);
        bytes.Position = 0;
        using var read = new EmailDocumentReader().Read(bytes);
        Assert.Contains("<b>rich</b>", read.Document.Body.Html);
        Assert.Contains("> Original rich body", read.Document.Body.Text);
        Assert.Contains("<img", original.Body.Html);
        Assert.Single(original.Headers);
    }

    [Fact]
    public void LargeRichQuoteFallsBackToBoundedTextWithoutBrokenMarkup() {
        var original = new EmailDocument();
        original.Body.Html = "<p><b>abc\ud83d\ude00more text</b></p>";
        var result = EmailHtmlComposer.Forward(original, new EmailAddress("me@example.test"), "Forward", new EmailHtmlCompositionOptions {
            Composition = new EmailCompositionOptions { MaxQuoteChars = 4 }
        });
        Assert.Contains("<blockquote><div>abc</div></blockquote>", result.Document.Body.Html);
        Assert.Equal("Forward\r\n\r\n> abc", result.Document.Body.Text);
        Assert.Empty(result.Document.Recipients);
        Assert.Null(result.Document.MessageMetadata.InReplyToId);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_COMPOSITION_TEXT_QUOTE_FALLBACK");
        var index = EmailIndexText.Create(result.Document, new EmailIndexTextOptions { PreferPlainText = false });
        Assert.Contains("abc", index.FullText);
        Assert.DoesNotContain("more text", index.FullText);
    }

    [Fact]
    public void NoQuotePolicyDoesNotParseOriginalAndStillEncodesAuthoredText() {
        var original = new EmailDocument();
        original.Body.Html = new string('x', 100);
        var result = EmailHtmlComposer.Reply(original, new EmailAddress("me@example.test"), "<script>authored</script>", new EmailHtmlCompositionOptions {
            MaxSourceChars = 1, Composition = new EmailCompositionOptions { QuoteOriginal = false }
        });
        Assert.DoesNotContain("<blockquote", result.Document.Body.Html);
        Assert.Contains("&lt;script&gt;authored&lt;/script&gt;", result.Document.Body.Html);
        Assert.Equal("<script>authored</script>", result.Document.Body.Text);
    }
}
