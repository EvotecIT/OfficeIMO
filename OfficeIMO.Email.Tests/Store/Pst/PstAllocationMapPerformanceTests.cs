#if NET8_0_OR_GREATER
using System.Diagnostics;
using Xunit.Abstractions;

namespace OfficeIMO.Email.Store.Tests;

[Trait("Category", "Performance")]
public sealed class PstAllocationMapPerformanceTests {
    private readonly ITestOutputHelper _output;

    public PstAllocationMapPerformanceTests(ITestOutputHelper output) {
        _output = output;
    }

    [Theory]
    [InlineData(32)]
    [InlineData(128)]
    [InlineData(512)]
    public void MeasuresWriterMapIoWithVerifiedAttachmentPayloads(int itemCount) {
        const int attachmentBytes = 128 * 1024;
        byte[] payload = Enumerable.Repeat((byte)0x5A, attachmentBytes).ToArray();
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-map-benchmark-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "benchmark.pst");
        try {
            long mapWrites;
            long outputBytes;
            var stopwatch = Stopwatch.StartNew();
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path,
                new EmailStorePstWriterOptions(failOnDataLoss: true, checkpointIntervalItems: int.MaxValue))) {
                PstAllocationMapTestSupport.CountingFileStream counted =
                    PstAllocationMapTestSupport.CountMapWrites(PstAllocationMapTestSupport.WriterFile(writer));
                string folder = writer.AddFolder("Benchmark");
                for (int index = 0; index < itemCount; index++) {
                    var document = new EmailDocument {
                        Subject = "Synthetic benchmark " + index,
                        MessageClass = "IPM.Note",
                        Date = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero)
                    };
                    document.Body.Text = "Body " + index;
                    document.Attachments.Add(new EmailAttachment {
                        FileName = "payload.bin", Length = attachmentBytes,
                        Content = payload
                    });
                    writer.AddItem(folder, document);
                }
                EmailStorePstWriteReport report = writer.Complete();
                Assert.Equal(itemCount, report.ItemCount);
                Assert.False(report.HasErrors);
                Assert.False(report.HasDataLoss);
                mapWrites = counted.WriteCalls;
                outputBytes = report.BytesWritten;
            }
            stopwatch.Stop();
            long writeMilliseconds = stopwatch.ElapsedMilliseconds;

            stopwatch.Restart();
            using EmailStoreSession session = EmailStoreSession.Open(path,
                new EmailStoreReaderOptions(retainAttachmentContent: true));
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (EmailStoreItemReference item in session.EnumerateItems()) {
                EmailDocument document = session.ReadItem(item).Document;
                Assert.NotNull(document.Subject);
                Assert.True(seen.Add(document.Subject!));
                int index = int.Parse(document.Subject!.Substring("Synthetic benchmark ".Length),
                    System.Globalization.CultureInfo.InvariantCulture);
                Assert.InRange(index, 0, itemCount - 1);
                Assert.Equal("Body " + index, document.Body.Text);
                EmailAttachment attachment = Assert.Single(document.Attachments);
                Assert.Equal("payload.bin", attachment.FileName);
                Assert.Equal(payload, attachment.Content);
            }
            Assert.Equal(itemCount, seen.Count);
            stopwatch.Stop();
            _output.WriteLine("PST_MAP_EVIDENCE items={0} attachmentBytes={1} outputBytes={2} mapWrites={3} writeMs={4} verifyMs={5}",
                itemCount, attachmentBytes, outputBytes, mapWrites, writeMilliseconds, stopwatch.ElapsedMilliseconds);
        } finally {
            try { if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
#endif
