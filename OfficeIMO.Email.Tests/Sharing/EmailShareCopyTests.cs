namespace OfficeIMO.Email.Tests;

public sealed class EmailShareCopyTests {
    [Theory]
    [InlineData("Resent-Bcc")]
    [InlineData("Resent-From")]
    [InlineData("Resent-To")]
    [InlineData("Resent-Cc")]
    [InlineData("Resent-Sender")]
    [InlineData("Resent-Date")]
    [InlineData("Resent-Message-ID")]
    public void ResentEnvelopeHeadersCannotBypassFieldSelectionInSavedArtifact(string headerName) {
        var source = new EmailDocument();
        source.Body.Text = "Reviewed text";
        source.Headers.Add(new EmailHeader(headerName, "private@example.test"));
        var options = new EmailShareCopyOptions();
        options.RetainedHeaderNames.Add(headerName);
        options.AddressReplacements.Add("private@example.test", null);
        var copy = EmailShareCopy.Create(source, options);
        using var output = new MemoryStream();
        Assert.False(new EmailDocumentWriter().Write(copy.Document, output).HasErrors);
        string artifact = Encoding.UTF8.GetString(output.ToArray());
        Assert.DoesNotContain("private@example.test", artifact);
        Assert.Empty(copy.Document.Headers);
        Assert.Contains(copy.Changes, change => change.Path == "headers/0" && change.Action == "omitted");
        Assert.Equal("private@example.test", Assert.Single(source.Headers).Value);
    }
    [Fact]
    public void AggregateAttachmentBudgetAndRetainedTextBoundsApplyToSelectedFields() {
        var source = new EmailDocument { From = new EmailAddress("long-address@example.test") };
        source.Attachments.Add(new EmailAttachment { Content = new byte[4] });
        source.Attachments.Add(new EmailAttachment { Content = new byte[4] });
        var options = new EmailShareCopyOptions { MaxAttachmentBytes = 4, MaxTotalAttachmentBytes = 7 };
        options.AttachmentIndexes.Add(0);
        options.AttachmentIndexes.Add(1);
        Assert.Throws<EmailLimitExceededException>(() => EmailShareCopy.Create(source, options));
        options.AttachmentIndexes.Clear();
        options.RetainedFields = EmailShareFields.Author;
        options.MaxTextChars = 5;
        Assert.Throws<EmailLimitExceededException>(() => EmailShareCopy.Create(source, options));
        options.AddressReplacements.Add("long-address@example.test", "a@b.c");
        Assert.Equal("a@b.c", EmailShareCopy.Create(source, options).Document.From!.Address);
    }
    [Fact]
    public void UnknownLengthAttachmentReadIsBoundedAndCancellationClosesItsSource() {
        var source = new EmailDocument();
        var content = new ObservedSource();
        source.Attachments.Add(new EmailAttachment { ContentSource = content });
        var options = new EmailShareCopyOptions { MaxAttachmentBytes = 10 };
        options.AttachmentIndexes.Add(0);
        Assert.Throws<EmailLimitExceededException>(() => EmailShareCopy.Create(source, options));
        Assert.Equal(11, content.Stream!.BytesRead);
        Assert.True(content.Stream.Disposed);
        using var cancellation = new CancellationTokenSource();
        content.Cancel = cancellation;
        options.MaxAttachmentBytes = 30;
        Assert.Throws<OperationCanceledException>(() => EmailShareCopy.Create(source, options, cancellation.Token));
        Assert.True(content.Stream!.Disposed);
    }

    [Fact]
    public void ExplicitConcealedHtmlCleanupUsesSharedSelectionOwnerWithoutChangingSource() {
        var source = new EmailDocument();
        source.Body.Html = "<p>Public</p><p style='display:none'>Concealed private text</p>";
        var inspection = OfficeIMO.Html.HtmlContentSafety.Inspect(source.Body.Html);
        var selection = new OfficeIMO.ContentSafety.OfficeContentCleanupSelection(inspection.Findings
            .Where(finding => finding.Kind == OfficeIMO.ContentSafety.OfficeContentConcealmentKind.HiddenByProperty).Select(finding => finding.Id));
        Assert.NotEmpty(selection.FindingIds);
        var result = EmailHtmlShareCopy.Create(source, cleanupSelection: selection);
        Assert.Contains("Public", result.Copy.Document.Body.Text);
        Assert.DoesNotContain("Concealed private text", result.Copy.Document.Body.Text);
        Assert.True(result.RemovedConcealedFindingCount > 0);
        Assert.Contains("Concealed private text", source.Body.Html);
    }
    [Fact]
    public void DefaultCopySerializesOnlyPlainBodyWithoutHiddenOriginalEnvelopeOrIntegrity() {
        const string text = "From: Secret Person <secret@example.test>\r\nTo: hidden@example.test\r\nBcc: private@example.test\r\n" +
            "Subject: Secret subject\r\nDKIM-Signature: original-signature\r\nReceived: confidential-route\r\nMessage-ID: <secret@example.test>\r\n" +
            "Content-Type: text/plain; charset=utf-8\r\n\r\nPublic body";
        using var read = new EmailDocumentReader().Read(new MemoryStream(Encoding.UTF8.GetBytes(text)));
        var result = EmailShareCopy.Create(read.Document);
        Assert.Equal("Public body", result.Document.Body.Text);
        Assert.Null(result.Document.From);
        Assert.Null(result.Document.Subject);
        Assert.Null(result.Document.RawSource);
        Assert.Empty(result.Document.Recipients);
        Assert.Empty(result.Document.Headers);
        Assert.Empty(result.Document.MapiProperties);
        Assert.Empty(result.Document.TnefAttributes);
        Assert.Empty(result.Document.Properties);
        Assert.False(result.Document.Protection.IsProtected);
        Assert.Contains(result.Diagnostics, value => value.Code == "EMAIL_SHARE_INTEGRITY_REMOVED");
        using var output = new MemoryStream();
        Assert.False(new EmailDocumentWriter().Write(result.Document, output).HasErrors);
        string artifact = Encoding.UTF8.GetString(output.ToArray());
        Assert.DoesNotContain("secret@example.test", artifact);
        Assert.DoesNotContain("confidential-route", artifact);
        Assert.DoesNotContain("original-signature", artifact);
        output.Position = 0;
        using var reopened = new EmailDocumentReader().Read(output);
        Assert.Equal("Public body", reopened.Document.Body.Text);
        Assert.NotEmpty(read.Document.Headers);
    }

    [Fact]
    public void ExplicitFieldPolicyReplacesAddressesAndNamesWithoutEnvelopeHeaderBypasses() {
        var source = new EmailDocument { From = new EmailAddress("secret@example.test", "Secret", "Raw-secret"), Subject = "Subject" };
        source.Body.Text = "Private body";
        source.Headers.Add(new EmailHeader("From", "bypass@example.test"));
        source.Headers.Add(new EmailHeader("DKIM-Signature", "signature"));
        source.Headers.Add(new EmailHeader("X-Category", "Review"));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("recipient@example.test", "Display")));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.Bcc, new EmailAddress("bcc@example.test")));
        byte[] payload = { 1, 2, 3 };
        source.Attachments.Add(new EmailAttachment { FileName = "secret.txt", Content = payload, ContentId = "private-cid", LinkedPath = "private-path" });
        var options = new EmailShareCopyOptions { RetainedFields = EmailShareFields.Subject | EmailShareFields.Author | EmailShareFields.Recipients, ReplacementBodyText = "Reviewed body" };
        options.AddressReplacements.Add("SECRET@example.test", "author@example.test");
        foreach (string header in new[] { "From", "DKIM-Signature", "X-Category" }) options.RetainedHeaderNames.Add(header);
        options.AttachmentIndexes.Add(0);
        options.AttachmentNames.Add(0, "shared.txt");
        var result = EmailShareCopy.Create(source, options);
        Assert.Equal("author@example.test", result.Document.From!.Address);
        Assert.Null(result.Document.From.DisplayName);
        Assert.Null(result.Document.From.RawValue);
        Assert.Single(result.Document.Recipients);
        Assert.Null(result.Document.Recipients[0].Address.DisplayName);
        Assert.Equal("X-Category", Assert.Single(result.Document.Headers).Name);
        var attachment = Assert.Single(result.Document.Attachments);
        Assert.Equal("shared.txt", attachment.FileName);
        Assert.Null(attachment.ContentId);
        Assert.Null(attachment.LinkedPath);
        payload[0] = 9;
        Assert.Equal(new byte[] { 1, 2, 3 }, attachment.Content);
        Assert.Contains(result.Changes, value => value.Path == "from" && value.Action == "replaced");
        Assert.Contains(result.Changes, value => value.Path == "recipients/1" && value.Action == "omitted");
        Assert.Contains(result.Diagnostics, value => value.Code == "EMAIL_SHARE_PAYLOAD_RETAINED");
        Assert.Equal("Private body", source.Body.Text);
        Assert.Equal("secret.txt", source.Attachments[0].FileName);
    }

    [Fact]
    public void SelectionRefusesUnavailableLinkedAndEmbeddedContentAndEnforcesDecodedBudgets() {
        var source = new EmailDocument();
        source.Attachments.Add(new EmailAttachment { LinkedPath = "must-not-open" });
        source.Attachments.Add(new EmailAttachment { EmbeddedDocument = new EmailDocument() });
        source.Attachments.Add(new EmailAttachment { Content = new byte[4] });
        var options = new EmailShareCopyOptions { MaxAttachmentBytes = 3 };
        options.AttachmentIndexes.Add(0);
        Assert.Throws<InvalidDataException>(() => EmailShareCopy.Create(source, options));
        options.AttachmentIndexes.Clear(); options.AttachmentIndexes.Add(1);
        Assert.Throws<ArgumentException>(() => EmailShareCopy.Create(source, options));
        options.AttachmentIndexes.Clear(); options.AttachmentIndexes.Add(2);
        Assert.Throws<EmailLimitExceededException>(() => EmailShareCopy.Create(source, options));
    }

    [Fact]
    public void HtmlBridgeProjectsSelectedTextAndPreservesCallerPolicy() {
        var source = new EmailDocument { Subject = "Private subject" };
        source.Body.Html = "<p>Public text</p><blockquote>Private history</blockquote><script>hidden-script</script>";
        var options = new EmailShareCopyOptions { ReplacementSubject = "Shared" };
        var result = EmailHtmlShareCopy.Create(source, options, new EmailIndexTextOptions { ExcludeQuotes = true });
        Assert.Equal("Public text\n", result.Copy.Document.Body.Text);
        Assert.Null(result.Copy.Document.Body.Html);
        Assert.Equal(EmailBodySourceKind.Html, result.BodySourceKind);
        Assert.DoesNotContain("hidden-script", result.Copy.Document.Body.Text);
        Assert.Null(options.ReplacementBodyText);
        Assert.Equal("Shared", result.Copy.Document.Subject);
    }

    private sealed class ObservedSource : IEmailContentSource {
        public long? Length => null;
        public CancellationTokenSource? Cancel { get; set; }
        public ObservedStream? Stream { get; private set; }
        public Stream OpenRead() => Stream = new ObservedStream(Cancel);
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) => Task.FromResult(OpenRead());
    }
    private sealed class ObservedStream : MemoryStream {
        private readonly CancellationTokenSource? _cancel;
        public ObservedStream(CancellationTokenSource? cancel) : base(new byte[20]) { _cancel = cancel; }
        public int BytesRead { get; private set; }
        public bool Disposed { get; private set; }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count); BytesRead += read; _cancel?.Cancel(); return read;
        }
        protected override void Dispose(bool disposing) { Disposed = true; base.Dispose(disposing); }
    }
}
