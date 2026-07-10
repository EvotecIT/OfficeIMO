#if NET8_0_OR_GREATER
using OfficeIMO.Email;
using System.Diagnostics;
using System.Globalization;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Email.Tests;

public sealed class EmailPerformanceEvidenceTests {
    private readonly ITestOutputHelper _output;

    public EmailPerformanceEvidenceTests(ITestOutputHelper output) {
        _output = output;
    }

    [Fact]
    public void MegabyteMimeRead_StaysInsideLinearAllocationAndTimeEnvelopes() {
        var document = new EmailDocument {
            Format = EmailFileFormat.Eml,
            Subject = "large-body"
        };
        document.Body.Text = new string('x', 1024 * 1024);
        byte[] source = new EmailDocumentWriter().WriteToBytes(document);
        var reader = new EmailDocumentReader(new EmailReaderOptions(maxInputBytes: 8L * 1024L * 1024L));
        reader.Read(Encoding.ASCII.GetBytes("Subject: warmup\r\n\r\nwarmup"));

        long before = GC.GetAllocatedBytesForCurrentThread();
        var stopwatch = Stopwatch.StartNew();
        EmailReadResult result = reader.Read(source);
        stopwatch.Stop();
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.Equal(1024 * 1024, result.Document.Body.Text!.Length);
        Assert.True(allocated <= (source.LongLength * 16L) + (16L * 1024L * 1024L),
            $"Allocated {allocated:N0} bytes for a {source.LongLength:N0}-byte MIME input.");
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(10),
            $"Reading a {source.LongLength:N0}-byte MIME input took {stopwatch.Elapsed}.");
        _output.WriteLine("MIME bytes: {0:N0}; allocated bytes: {1:N0}; elapsed: {2}",
            source.LongLength, allocated, stopwatch.Elapsed);
    }

    [Fact]
    public void FiveHundredMessageMboxRead_StaysInsideAggregateAllocationAndTimeEnvelopes() {
        var mailbox = new EmailMailbox();
        for (int index = 0; index < 500; index++) {
            var document = new EmailDocument {
                Format = EmailFileFormat.Eml,
                Subject = string.Concat("message-", index.ToString(CultureInfo.InvariantCulture)),
                From = new EmailAddress("sender@example.test"),
                Date = new DateTimeOffset(2026, 7, 10, 12, 0, 0, TimeSpan.Zero)
            };
            document.Body.Text = string.Concat("body-", index.ToString(CultureInfo.InvariantCulture));
            mailbox.Messages.Add(new EmailMailboxEntry(document) { EnvelopeSender = "sender@example.test" });
        }
        byte[] source = new EmailMailboxWriter().WriteToBytes(mailbox);
        var reader = new EmailMailboxReader(new EmailMailboxReaderOptions(
            new EmailReaderOptions(maxInputBytes: 16L * 1024L * 1024L, includeAttachmentContent: false),
            maxMessageCount: 500));

        long before = GC.GetAllocatedBytesForCurrentThread();
        var stopwatch = Stopwatch.StartNew();
        EmailMailboxReadResult result = reader.Read(source);
        stopwatch.Stop();
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.Equal(500, result.Mailbox.Messages.Count);
        Assert.True(allocated <= (source.LongLength * 64L) + (32L * 1024L * 1024L),
            $"Allocated {allocated:N0} bytes for a {source.LongLength:N0}-byte mbox input.");
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(10),
            $"Reading a 500-message mbox took {stopwatch.Elapsed}.");
        _output.WriteLine("Mbox bytes: {0:N0}; allocated bytes: {1:N0}; elapsed: {2}",
            source.LongLength, allocated, stopwatch.Elapsed);
    }

    [Fact]
    public void MegabyteMsgRead_UsesOneBoundedCompoundOpenWithinAllocationAndTimeEnvelopes() {
        var document = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "large-msg"
        };
        document.Body.Text = "body";
        document.Attachments.Add(new EmailAttachment {
            FileName = "payload.bin",
            ContentType = "application/octet-stream",
            Content = new byte[1024 * 1024],
            Length = 1024 * 1024
        });
        byte[] source = new EmailDocumentWriter().WriteToBytes(document, EmailFileFormat.OutlookMsg);
        var reader = new EmailDocumentReader(new EmailReaderOptions(maxInputBytes: 8L * 1024L * 1024L));

        long before = GC.GetAllocatedBytesForCurrentThread();
        var stopwatch = Stopwatch.StartNew();
        EmailReadResult result = reader.Read(source);
        stopwatch.Stop();
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.Equal(1024 * 1024, Assert.Single(result.Document.Attachments).Content!.Length);
        Assert.True(allocated <= (source.LongLength * 32L) + (32L * 1024L * 1024L),
            $"Allocated {allocated:N0} bytes for a {source.LongLength:N0}-byte MSG input.");
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(10),
            $"Reading a {source.LongLength:N0}-byte MSG input took {stopwatch.Elapsed}.");
        _output.WriteLine("MSG bytes: {0:N0}; allocated bytes: {1:N0}; elapsed: {2}",
            source.LongLength, allocated, stopwatch.Elapsed);
    }
}
#endif
