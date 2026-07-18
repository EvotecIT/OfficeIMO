using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailMailboxTests {
    [Fact]
    public void ReadWritePreservesAnUnparseableRawEnvelopeLine() {
        const string rawFromLine = "From vendor@example.test Vendor-Date-Token";
        byte[] source = Encoding.ASCII.GetBytes(rawFromLine +
            "\nDate: Fri, 17 Jul 2026 09:00:00 +0000\nSubject: Vendor envelope\n\nBody\n");

        EmailMailbox mailbox = EmailMailbox.Load(source);
        EmailMailboxEntry entry = Assert.Single(mailbox.Messages);
        byte[] serialized = mailbox.ToBytes();
        EmailMailboxEntry reparsed = Assert.Single(EmailMailbox.Load(serialized).Messages);

        Assert.Equal(rawFromLine, entry.RawFromLine);
        Assert.Null(entry.EnvelopeDate);
        Assert.StartsWith(rawFromLine + "\n", Encoding.ASCII.GetString(serialized),
            StringComparison.Ordinal);
        Assert.Equal(rawFromLine, reparsed.RawFromLine);
        Assert.Null(reparsed.EnvelopeDate);
    }

    [Theory]
    [InlineData("From attacker@example.test\nFrom injected@example.test")]
    [InlineData("From attacker@example.test\u007Ftrailer")]
    public void WriterRegeneratesAnUnsafeRawEnvelopeLineWithoutInjectingASeparator(
        string unsafeRawFromLine) {
        var mailbox = new EmailMailbox();
        mailbox.Messages.Add(new EmailMailboxEntry(new EmailDocument {
            From = new EmailAddress("sender@example.test"),
            Date = new DateTimeOffset(2026, 7, 17, 9, 0, 0, TimeSpan.Zero),
            Subject = "Safe envelope"
        }) {
            RawFromLine = unsafeRawFromLine
        });
        using var output = new MemoryStream();

        EmailWriteResult result = new EmailMailboxWriter().Write(mailbox, output);
        string serialized = Encoding.ASCII.GetString(output.ToArray());

        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_MBOX_RAW_ENVELOPE_INVALID" &&
            diagnostic.Severity == EmailDiagnosticSeverity.Warning);
        Assert.StartsWith("From sender@example.test Fri Jul 17 09:00:00 2026\n",
            serialized, StringComparison.Ordinal);
        Assert.DoesNotContain(unsafeRawFromLine, serialized, StringComparison.Ordinal);
    }

    [Fact]
    public void MailboxToStreamMatchesToBytesAndStartsAtBeginning() {
        var mailbox = new EmailMailbox();

        using MemoryStream stream = mailbox.ToStream();

        Assert.Equal(0, stream.Position);
        Assert.Equal(mailbox.ToBytes(), stream.ToArray());
        Assert.True(stream.CanWrite);
    }

    [Fact]
    public void ReadsAndWritesMboxrdWithEnvelopeAndFromLineEscaping() {
        const string source = "From first@example.com Fri Jul 10 12:00:00 2026\n" +
            "From: first@example.com\nSubject: First\nContent-Type: text/plain; charset=utf-8\n\n" +
            ">From escaped separator\n>>From originally quoted\n" +
            "From second@example.com Fri Jul 10 13:00:00 2026\n" +
            "From: second@example.com\nSubject: Second\nContent-Type: text/plain; charset=utf-8\n\nsecond body\n";
        var messageOptions = new EmailReaderOptions(preserveRawSource: true);
        var reader = new EmailMailboxReader(new EmailMailboxReaderOptions(messageOptions));

        EmailMailboxReadResult result = reader.Read(Encoding.UTF8.GetBytes(source));

        Assert.Equal(2, result.Mailbox.Messages.Count);
        Assert.Equal("first@example.com", result.Mailbox.Messages[0].EnvelopeSender);
        Assert.Contains("From escaped separator", result.Mailbox.Messages[0].Document.Body.Text);
        Assert.Contains(">From originally quoted", result.Mailbox.Messages[0].Document.Body.Text);
        Assert.Equal("Second", result.Mailbox.Messages[1].Document.Subject);

        var writerOptions = new EmailMailboxWriterOptions(
            new EmailWriterOptions(usePreservedRawSource: true), MboxVariant.Mboxrd);
        byte[] rewritten = new EmailMailboxWriter(writerOptions).ToBytes(result.Mailbox);
        EmailMailboxReadResult reparsed = reader.Read(rewritten);
        Assert.Equal(2, reparsed.Mailbox.Messages.Count);
        Assert.Contains("From escaped separator", reparsed.Mailbox.Messages[0].Document.Body.Text);
        Assert.Contains(">From originally quoted", reparsed.Mailbox.Messages[0].Document.Body.Text);
    }

    [Fact]
    public async Task AsyncMailboxReadLeavesStreamOpenAndEnforcesMessageLimit() {
        byte[] source = Encoding.ASCII.GetBytes(
            "From a@example.com Fri Jul 10 12:00:00 2026\nSubject: A\n\nA\n" +
            "From b@example.com Fri Jul 10 13:00:00 2026\nSubject: B\n\nB\n");
        MemoryStream stream = new MemoryStream(source);
        stream.Position = 7;

        EmailMailboxReadResult result = await new EmailMailboxReader().ReadAsync(stream);

        Assert.Equal(2, result.Mailbox.Messages.Count);
        Assert.True(stream.CanRead);
        Assert.Equal(7, stream.Position);
        Assert.Throws<EmailLimitExceededException>(() => new EmailMailboxReader(
            new EmailMailboxReaderOptions(maxMessageCount: 1)).Read(source));
        stream.Dispose();
    }

    [Fact]
    public void StopsEnvelopeDiscoveryAtTheFirstMessagePastTheConfiguredLimit() {
        byte[] source = Encoding.ASCII.GetBytes(
            "From a@example.com Fri Jul 10 12:00:00 2026\nSubject: A\n\nA\n" +
            "From b@example.com Fri Jul 10 13:00:00 2026\nSubject: B\n\nB\n" +
            "From c@example.com Fri Jul 10 14:00:00 2026\nSubject: C\n\nC\n");

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailMailboxReader(new EmailMailboxReaderOptions(maxMessageCount: 1)).Read(source));

        Assert.Equal(nameof(EmailMailboxReaderOptions.MaxMessageCount), exception.LimitName);
        Assert.Equal(2, exception.ActualValue);
    }

    [Fact]
    public void BoundsAggregateMailboxOutputAcrossIndividuallyValidMessages() {
        var messageOptions = new EmailWriterOptions(maxOutputBytes: 220);
        var first = new EmailDocument { Subject = "first" };
        first.Body.Text = new string('a', 40);
        var second = new EmailDocument { Subject = "second" };
        second.Body.Text = new string('b', 40);
        Assert.True(new EmailDocumentWriter(messageOptions).ToBytes(first).Length < 220);
        Assert.True(new EmailDocumentWriter(messageOptions).ToBytes(second).Length < 220);
        var mailbox = new EmailMailbox();
        mailbox.Messages.Add(new EmailMailboxEntry(first));
        mailbox.Messages.Add(new EmailMailboxEntry(second));

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailMailboxWriter(new EmailMailboxWriterOptions(messageOptions)).ToBytes(mailbox));

        Assert.Equal(nameof(EmailWriterOptions.MaxOutputBytes), exception.LimitName);
        Assert.Equal(220, exception.MaximumValue);
    }

    [Fact]
    public void StreamsCumulativeMailboxOutputBeyondThePerMessageLimit() {
        var messageOptions = new EmailWriterOptions(maxOutputBytes: 220);
        var first = new EmailDocument { Subject = "first" };
        first.Body.Text = new string('a', 40);
        var second = new EmailDocument { Subject = "second" };
        second.Body.Text = new string('b', 40);
        var entries = new[] { new EmailMailboxEntry(first), new EmailMailboxEntry(second) };
        using var output = new MemoryStream();

        EmailWriteResult result = new EmailMailboxWriter(new EmailMailboxWriterOptions(messageOptions))
            .WriteEntries(entries, output);

        Assert.True(result.BytesWritten > messageOptions.MaxOutputBytes);
        Assert.Equal(output.Length, result.BytesWritten);
    }

    [Fact]
    public void SingleDocumentReaderDirectsMboxToAggregateApi() {
        byte[] source = Encoding.ASCII.GetBytes("From a@example.com Fri Jul 10 12:00:00 2026\nSubject: A\n\nA\n");

        EmailReadResult result = new EmailDocumentReader().Read(source);

        Assert.Equal(EmailFileFormat.Mbox, result.Document.Format);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_MBOX_REQUIRES_MAILBOX_READER");
    }

    [Fact]
    public void MailboxWriterReturnsPerMessageDiagnostics() {
        var document = new EmailDocument { Subject = "lossy attachment" };
        document.Attachments.Add(new EmailAttachment { FileName = "missing.bin", Length = 12 });
        var mailbox = new EmailMailbox();
        mailbox.Messages.Add(new EmailMailboxEntry(document));

        using var output = new MemoryStream();
        EmailWriteResult result = new EmailMailboxWriter().Write(mailbox, output);

        Assert.Contains(result.Diagnostics,
            diagnostic => diagnostic.Code == "EMAIL_ATTACHMENT_CONTENT_UNAVAILABLE");
    }

    [Fact]
    public void EnumeratesAndWritesMailboxEntriesWithoutAggregateMaterialization() {
        byte[] source = Encoding.ASCII.GetBytes(
            "From a@example.com Fri Jul 10 12:00:00 2026\nSubject: A\n\nA\n" +
            "From b@example.com Fri Jul 10 13:00:00 2026\nSubject: B\n\nB\n");
        using var input = new MemoryStream(source);
        input.Position = 3;
        var reader = new EmailMailboxReader();

        EmailMailboxEntryReadResult[] entries = reader.ReadEntries(input).ToArray();
        using var output = new MemoryStream();
        EmailWriteResult write = new EmailMailboxWriter().WriteEntries(entries.Select(item => item.Entry), output);
        EmailMailboxEntryReadResult[] roundTrip = reader.ReadEntries(new MemoryStream(output.ToArray())).ToArray();

        Assert.Equal(3, input.Position);
        Assert.Equal(new[] { "A", "B" }, entries.Select(item => item.Entry.Document.Subject));
        Assert.Equal(2, roundTrip.Length);
        Assert.Equal(output.Length, write.BytesWritten);
    }

    [Fact]
    public void StreamingReaderAppliesMailboxAndPerMessageLimitsSeparately() {
        byte[] source = Encoding.ASCII.GetBytes(
            "From a@example.com Fri Jul 10 12:00:00 2026\nSubject: A\n\n1234567890\n" +
            "From b@example.com Fri Jul 10 13:00:00 2026\nSubject: B\n\n1234567890\n");
        var options = new EmailMailboxReaderOptions(maxMailboxBytes: 200,
            messageOptions: new EmailReaderOptions(maxInputBytes: 80));

        EmailMailboxEntryReadResult[] entries = new EmailMailboxReader(options)
            .ReadEntries(new MemoryStream(source)).ToArray();

        Assert.Equal(2, entries.Length);
    }

    [Fact]
    public async Task AggregateReadersReportTheAggregateMailboxLimit() {
        byte[] source = Encoding.ASCII.GetBytes(
            "From a@example.com Fri Jul 10 12:00:00 2026\nSubject: A\n\n1234567890\n");
        var options = new EmailMailboxReaderOptions(maxMailboxBytes: 20);
        var reader = new EmailMailboxReader(options);

        EmailLimitExceededException synchronous = Assert.Throws<EmailLimitExceededException>(() =>
            reader.Read(new MemoryStream(source)));
        EmailLimitExceededException asynchronous = await Assert.ThrowsAsync<EmailLimitExceededException>(() =>
            reader.ReadAsync(new MemoryStream(source)));

        Assert.Equal(nameof(EmailMailboxReaderOptions.MaxMailboxBytes), synchronous.LimitName);
        Assert.Equal(nameof(EmailMailboxReaderOptions.MaxMailboxBytes), asynchronous.LimitName);
    }

    [Fact]
    public void AcceptsUtf8BomAndRecoversUnescapedBodyFromLinesAsEntries() {
        byte[] source = new byte[] { 0xEF, 0xBB, 0xBF }.Concat(Encoding.ASCII.GetBytes(
            "From -\r\nSubject: First\r\n\r\nFrom Russia with love\r\n" +
            "From -\r\nSubject: Second\r\n\r\nbody\r\n")).ToArray();
        var reader = new EmailMailboxReader();

        EmailMailboxReadResult aggregate = reader.Read(source);
        EmailMailboxEntryReadResult[] streamed = reader.ReadEntries(new MemoryStream(source)).ToArray();

        Assert.Equal(3, aggregate.Mailbox.Messages.Count);
        Assert.Equal(3, streamed.Length);
        Assert.Equal("First", aggregate.Mailbox.Messages[0].Document.Subject);
        Assert.Equal("Second", aggregate.Mailbox.Messages[2].Document.Subject);
        Assert.Contains(aggregate.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_MBOX_MESSAGE_HEADERS_MISSING");
        Assert.DoesNotContain(streamed.SelectMany(entry => entry.Diagnostics), diagnostic =>
            diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    [Fact]
    public void PreservesHeaderlessMboxEntryBody() {
        byte[] source = Encoding.ASCII.GetBytes(
            "From sender@example.com Fri Jul 10 12:00:00 2026\nplain body\n");
        var reader = new EmailMailboxReader();

        EmailMailboxEntry aggregate = Assert.Single(reader.Read(source).Mailbox.Messages);
        EmailMailboxEntryReadResult streamed = Assert.Single(
            reader.ReadEntries(new MemoryStream(source)).ToArray());

        Assert.Equal("plain body", aggregate.Document.Body.Text!.Trim());
        Assert.Equal("plain body", streamed.Entry.Document.Body.Text!.Trim());
        Assert.Contains(streamed.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_MBOX_MESSAGE_HEADERS_MISSING");
    }

    [Fact]
    public async Task AggregateStreamReadersPreservePerMessageLimitFailures() {
        byte[] source = Encoding.ASCII.GetBytes(
            "From sender@example.com Fri Jul 10 12:00:00 2026\nSubject: oversized\n\n" +
            new string('x', 100) + "\n");
        var options = new EmailMailboxReaderOptions(maxMailboxBytes: source.Length + 10,
            messageOptions: new EmailReaderOptions(maxInputBytes: 64));
        var reader = new EmailMailboxReader(options);

        EmailLimitExceededException synchronous = Assert.Throws<EmailLimitExceededException>(() =>
            reader.Read(new MemoryStream(source)));
        EmailLimitExceededException asynchronous = await Assert.ThrowsAsync<EmailLimitExceededException>(() =>
            reader.ReadAsync(new MemoryStream(source)));

        Assert.Equal(nameof(EmailReaderOptions.MaxInputBytes), synchronous.LimitName);
        Assert.Equal(nameof(EmailReaderOptions.MaxInputBytes), asynchronous.LimitName);
    }

    [Fact]
    public void StopsReadingAnOversizedUnterminatedMboxLineAtTheMessageLimit() {
        byte[] source = Encoding.ASCII.GetBytes(
            "From sender@example.com Fri Jul 10 12:00:00 2026\n" + new string('x', 10000));
        var options = new EmailMailboxReaderOptions(maxMailboxBytes: source.Length + 10,
            messageOptions: new EmailReaderOptions(maxInputBytes: 64));
        using var stream = new GuardedMemoryStream(source, maximumReads: 200);

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailMailboxReader(options).ReadEntries(stream).ToArray());

        Assert.Equal(nameof(EmailReaderOptions.MaxInputBytes), exception.LimitName);
        Assert.True(stream.BytesRead < 200);
    }

    [Fact]
    public void StopsReadingAnOversizedFromPrefixedMboxLineAtTheMessageLimit() {
        byte[] source = Encoding.ASCII.GetBytes(
            "From sender@example.com Fri Jul 10 12:00:00 2026\nFrom " + new string('x', 10000));
        var options = new EmailMailboxReaderOptions(maxMailboxBytes: source.Length + 10,
            messageOptions: new EmailReaderOptions(maxInputBytes: 64));
        using var stream = new GuardedMemoryStream(source, maximumReads: 200);

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailMailboxReader(options).ReadEntries(stream).ToArray());

        Assert.Equal(nameof(EmailReaderOptions.MaxInputBytes), exception.LimitName);
        Assert.True(stream.BytesRead < 200);
    }

    [Fact]
    public void StopsReadingAnOversizedUnterminatedMboxLineAtTheMailboxLimit() {
        byte[] source = Encoding.ASCII.GetBytes(new string('x', 10000));
        var options = new EmailMailboxReaderOptions(maxMailboxBytes: 64);
        using var stream = new GuardedMemoryStream(source, maximumReads: 200);

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            new EmailMailboxReader(options).ReadEntries(stream).ToArray());

        Assert.Equal(nameof(EmailMailboxReaderOptions.MaxMailboxBytes), exception.LimitName);
        Assert.True(stream.BytesRead < 200);
    }

    private sealed class GuardedMemoryStream : MemoryStream {
        private readonly int _maximumReads;

        internal GuardedMemoryStream(byte[] source, int maximumReads) : base(source) {
            _maximumReads = maximumReads;
        }

        internal int BytesRead { get; private set; }

        public override int ReadByte() {
            if (BytesRead >= _maximumReads) throw new InvalidOperationException("The reader did not stop at its limit.");
            int value = base.ReadByte();
            if (value >= 0) BytesRead++;
            return value;
        }
    }
}
