namespace OfficeIMO.Email.Tests;

public sealed class EmailCompositionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RepliesRetainSingleInReplyToAncestorWhenReferencesIsAbsent(bool retainedHeader) {
        var original = new EmailDocument { From = new EmailAddress("sender@example.test"), MessageId = "<parent@example.test>" };
        if (retainedHeader) original.Headers.Add(new EmailHeader("In-Reply-To", "<root@example.test>"));
        else original.MessageMetadata.InReplyToId = "<root@example.test>";
        var reply = EmailComposer.Reply(original, new EmailAddress("me@example.test"), "Hi", new EmailCompositionOptions { QuoteOriginal = false });
        Assert.Equal("<root@example.test> <parent@example.test>", reply.Document.MessageMetadata.InternetReferences);
    }
    [Fact]
    public void ReplyAllUsesReplyToDeduplicatesAndExcludesOwnAliasesAndBcc() {
        var source = new EmailDocument { From = new EmailAddress("author@example.test"), Subject = "Topic", MessageId = "<parent@example.test>" };
        source.Body.Text = "Original body";
        source.Headers.Add(new EmailHeader("DKIM-Signature", "old"));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.ReplyTo, new EmailAddress("reply@example.test", "Reply")));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("ME@example.test")));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("alias@example.test")));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("reply@example.test")));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("other@example.test")));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.Cc, new EmailAddress("OTHER@example.test")));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.Cc, new EmailAddress("cc@example.test")));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.Bcc, new EmailAddress("hidden@example.test")));
        source.MessageMetadata.InternetReferences = "<root@example.test> <root@example.test>";
        source.Attachments.Add(new EmailAttachment { Content = new byte[] { 1 } });
        var policy = new EmailCompositionOptions { Date = DateTimeOffset.Parse("2026-09-30T12:00:00Z") };
        policy.OwnAddresses.Add("alias@example.test");
        EmailCompositionResult result = EmailComposer.ReplyAll(source, new EmailAddress("me@example.test"), "Thanks", policy);
        Assert.Equal(new[] { "reply@example.test", "other@example.test", "cc@example.test" }, result.Document.Recipients.Select(item => item.Address.Address));
        Assert.Equal(new[] { EmailRecipientKind.To, EmailRecipientKind.To, EmailRecipientKind.Cc }, result.Document.Recipients.Select(item => item.Kind));
        Assert.Equal("<parent@example.test>", result.Document.MessageMetadata.InReplyToId);
        Assert.Equal("<root@example.test> <parent@example.test>", result.Document.MessageMetadata.InternetReferences);
        Assert.Equal("Re: Topic", result.Document.Subject);
        Assert.True(result.Document.MessageMetadata.IsDraft);
        Assert.Empty(result.Document.Headers);
        Assert.Empty(result.Document.Attachments);
        Assert.Null(result.Document.MessageId);
        Assert.Empty(result.Diagnostics);
        using EmailReadResult reread = new EmailDocumentReader().Read(result.Document.ToBytes());
        Assert.Equal("<parent@example.test>", reread.Document.MessageMetadata.InReplyToId);
        Assert.DoesNotContain(reread.Document.Recipients, item => item.Kind == EmailRecipientKind.Bcc);
        result.Document.Recipients[0].Address.DisplayName = "Changed";
        Assert.Equal("Reply", source.Recipients[0].Address.DisplayName);
        Assert.Equal("Original body", source.Body.Text);
        Assert.Single(source.Headers);
    }

    [Fact]
    public void ReplyAndForwardKeepDistinctRecipientAndThreadingPolicies() {
        var source = new EmailDocument { From = new EmailAddress("author@example.test"), Subject = "Re: Topic", MessageId = "<parent@example.test>" };
        source.Body.Text = "Body";
        EmailDocument reply = EmailComposer.Reply(source, new EmailAddress("me@example.test"), "Reply").Document;
        Assert.Equal("Re: Topic", reply.Subject);
        Assert.Equal("author@example.test", Assert.Single(reply.Recipients).Address.Address);
        EmailDocument forward = EmailComposer.Forward(source, new EmailAddress("me@example.test"), "Forward").Document;
        Assert.Equal("Fwd: Re: Topic", forward.Subject);
        Assert.Empty(forward.Recipients);
        Assert.Null(forward.MessageMetadata.InReplyToId);
        Assert.Contains("Forwarded message", forward.Body.Text);
    }

    [Fact]
    public void QuoteAndReferencesAreBoundedAndHtmlOnlyContentIsReported() {
        var source = new EmailDocument { From = new EmailAddress("author@example.test"), MessageId = "<parent@example.test>" };
        source.Body.Text = "A\U0001f600Z";
        source.MessageMetadata.InternetReferences = "<one@example.test> <two@example.test>";
        var result = EmailComposer.Reply(source, new EmailAddress("me@example.test"), "Hi", new EmailCompositionOptions { MaxQuoteChars = 2, MaxReferences = 1 });
        Assert.EndsWith("> A", result.Document.Body.Text);
        Assert.Equal("<parent@example.test>", result.Document.MessageMetadata.InternetReferences);
        Assert.Contains(result.Diagnostics, item => item.Code == "EMAIL_COMPOSITION_QUOTE_TRUNCATED");
        Assert.Contains(result.Diagnostics, item => item.Code == "EMAIL_COMPOSITION_REFERENCES_TRUNCATED");
        source.Body.Text = null;
        source.Body.Html = "<p>HTML</p>";
        Assert.Contains(EmailComposer.Reply(source, new EmailAddress("me@example.test"), "Hi").Diagnostics, item => item.Code == "EMAIL_COMPOSITION_PLAIN_BODY_UNAVAILABLE");
    }

    [Fact]
    public void UnresolvedExchangeReplyTargetRequiresExplicitResolutionAndNeverFallsBackSilently() {
        var source = new EmailDocument { From = new EmailAddress("author@example.test") };
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.ReplyTo, new EmailAddress("/o=organization/cn=user") { AddressType = "EX" }));
        var result = EmailComposer.Reply(source, new EmailAddress("me@example.test"), "Hi", new EmailCompositionOptions { QuoteOriginal = false });
        Assert.Empty(result.Document.Recipients);
        Assert.Contains(result.Diagnostics, item => item.Code == "EMAIL_COMPOSITION_UNRESOLVED_ADDRESS");
        Assert.Contains(result.Diagnostics, item => item.Code == "EMAIL_COMPOSITION_NO_RECIPIENTS");
        Assert.Throws<ArgumentException>(() => EmailComposer.Reply(source, new EmailAddress("me@example.test\r\nBcc: other@example.test"), "Hi"));
    }
}
