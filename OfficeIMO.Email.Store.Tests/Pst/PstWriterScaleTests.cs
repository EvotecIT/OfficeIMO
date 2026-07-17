#if NET8_0_OR_GREATER
using OfficeIMO.Email;
using System.Diagnostics;
using System.Globalization;
using Xunit.Abstractions;

namespace OfficeIMO.Email.Store.Tests;

public sealed class PstWriterScaleTests {
    private readonly ITestOutputHelper _output;

    public PstWriterScaleTests(ITestOutputHelper output) {
        _output = output;
    }

    [Fact]
    public void TwoThousandMessagePstUsesBoundedRetainedManagedMemory() {
        const int itemCount = 2_000;
        string path = Path.Combine(Path.GetTempPath(),
            string.Concat("officeimo-scale-", Guid.NewGuid().ToString("N"), ".pst"));
        try {
            var document = new EmailDocument {
                Subject = "Synthetic scale message",
                MessageClass = "IPM.Note",
                Date = new DateTimeOffset(2026, 7, 17, 0, 0, 0, TimeSpan.Zero)
            };
            document.Body.Text = "A small deterministic body proves high item cardinality without a private corpus.";
            using var writer = EmailStorePstWriter.Create(path,
                new EmailStorePstWriterOptions(
                    maxItemCount: itemCount,
                    checkpointIntervalItems: itemCount,
                    maxIndexRecordsInMemory: 256,
                    retainCheckpointOnDispose: false));
            string folder = writer.AddFolder("Scale");
            ForceCollection();
            long retainedBefore = GC.GetTotalMemory(forceFullCollection: false);
            var stopwatch = Stopwatch.StartNew();
            for (int index = 0; index < itemCount; index++) {
                writer.AddItem(folder, document);
            }
            ForceCollection();
            long retainedAfterItems = GC.GetTotalMemory(forceFullCollection: false);
            long retainedGrowth = Math.Max(0, retainedAfterItems - retainedBefore);
            EmailStorePstWriteReport report = writer.Complete();
            stopwatch.Stop();

            Assert.Equal(itemCount, report.ItemCount);
            Assert.True(retainedGrowth <= 64L * 1024L * 1024L,
                $"Retained managed memory grew by {retainedGrowth:N0} bytes for {itemCount:N0} items.");
            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(45),
                $"Creating {itemCount:N0} items took {stopwatch.Elapsed}.");
            using EmailStoreSession session = EmailStoreSession.Open(path,
                new EmailStoreReaderOptions(maxItemCount: itemCount));
            Assert.Equal(itemCount, session.EnumerateItems(
                new EmailStoreEnumerationOptions(maxItems: itemCount)).Count());
            _output.WriteLine(
                "PST items: {0:N0}; bytes: {1:N0}; retained growth: {2:N0}; elapsed: {3}",
                itemCount, report.BytesWritten, retainedGrowth, stopwatch.Elapsed);
        } finally {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    [Fact]
    public void HundredThousandVerificationMappingsRemainDiskBacked() {
        const int mappingCount = 100_000;
        string destination = Path.Combine(Path.GetTempPath(),
            string.Concat("officeimo-map-scale-", Guid.NewGuid().ToString("N"), ".pst"));
        ForceCollection();
        long retainedBefore = GC.GetTotalMemory(forceFullCollection: false);
        using var journal = new PstConversionMappingJournal(destination);
        for (int index = 0; index < mappingCount; index++) {
            string id = index.ToString("X8", CultureInfo.InvariantCulture);
            journal.Add(index + 1,
                new EmailStoreItemReference(string.Concat("source-", id), string.Concat("folder-", id),
                    isAssociated: (index & 1) != 0, isOrphaned: false),
                string.Concat("destination-folder-", id), string.Concat("destination-item-", id));
        }
        ForceCollection();
        long retainedAfter = GC.GetTotalMemory(forceFullCollection: false);
        long retainedGrowth = Math.Max(0, retainedAfter - retainedBefore);

        int read = 0;
        foreach (PstConversionItemMap mapping in journal.ReadAll()) {
            read++;
            if (read == 1) Assert.Equal("source-00000000", mapping.Source.Id);
            if (read == mappingCount) Assert.Equal("destination-item-0001869F", mapping.DestinationItemId);
        }

        Assert.Equal(mappingCount, journal.Count);
        Assert.Equal(mappingCount, read);
        Assert.True(journal.Length > mappingCount * 32L);
        Assert.True(retainedGrowth <= 32L * 1024L * 1024L,
            $"Retained managed memory grew by {retainedGrowth:N0} bytes for {mappingCount:N0} mappings.");
        _output.WriteLine("Mappings: {0:N0}; journal bytes: {1:N0}; retained growth: {2:N0}",
            mappingCount, journal.Length, retainedGrowth);
    }

    [Fact]
    public void HundredThousandSemanticDeduplicationDigestsRemainDiskBacked() {
        const int digestCount = 100_000;
        string destination = Path.Combine(Path.GetTempPath(),
            string.Concat("officeimo-dedup-scale-", Guid.NewGuid().ToString("N"), ".pst"));
        ForceCollection();
        long retainedBefore = GC.GetTotalMemory(forceFullCollection: false);
        using var index = new EmailSemanticDedupIndex(destination);
        for (int value = 0; value < digestCount; value++) {
            var digest = new byte[32];
            byte[] number = BitConverter.GetBytes(value);
            Buffer.BlockCopy(number, 0, digest, 0, number.Length);
            Assert.True(index.Add(digest));
        }
        ForceCollection();
        long retainedGrowth = Math.Max(0,
            GC.GetTotalMemory(forceFullCollection: false) - retainedBefore);
        var existing = new byte[32];
        Buffer.BlockCopy(BitConverter.GetBytes(digestCount / 2), 0, existing, 0, 4);

        Assert.Equal(digestCount, index.Count);
        Assert.True(index.Contains(existing));
        Assert.False(index.Add(existing));
        Assert.True(index.Length > digestCount * 33L);
        Assert.True(retainedGrowth <= 16L * 1024L * 1024L,
            $"Retained managed memory grew by {retainedGrowth:N0} bytes for {digestCount:N0} digests.");
        _output.WriteLine("Dedup digests: {0:N0}; index bytes: {1:N0}; retained growth: {2:N0}",
            digestCount, index.Length, retainedGrowth);
    }

    private static void ForceCollection() {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }
}
#endif
