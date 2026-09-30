using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests;

public sealed class PstWriterIndexCorrectnessTests {
    [Fact]
    public void Sorted_index_runs_preserve_every_message_when_reopened() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-index-" + Guid.NewGuid().ToString("N") + ".pst");
        try {
            string[] expected = Enumerable.Range(0, 12).Select(value => "Message " + value).ToArray();
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path,
                new EmailStorePstWriterOptions(maxIndexRecordsInMemory: 2, retainCheckpointOnDispose: false))) {
                string first = writer.AddFolder("First");
                string second = writer.AddFolder("Second");
                for (int index = 0; index < expected.Length; index++) {
                    writer.AddItem((index & 1) == 0 ? first : second, new EmailDocument { Subject = expected[index] });
                }
                Assert.Equal(expected.Length, writer.Complete().ItemCount);
            }
            using EmailStoreSession session = EmailStoreSession.Open(path);
            string[] actual = session.EnumerateItems().Select(reference => session.ReadItem(reference).Document.Subject!).ToArray();
            Assert.Equal(expected.OrderBy(value => value, StringComparer.Ordinal), actual.OrderBy(value => value, StringComparer.Ordinal));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Semantic_deduplication_growth_preserves_membership_and_duplicate_rejection() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-dedup-" + Guid.NewGuid().ToString("N") + ".pst");
        using var index = new EmailSemanticDedupIndex(path, initialCapacity: 16);
        byte[][] digests = Enumerable.Range(0, 32).Select(value => {
            var digest = new byte[32];
            digest[0] = (byte)value;
            return digest;
        }).ToArray();
        foreach (byte[] digest in digests) Assert.True(index.Add(digest));
        Assert.Equal(digests.Length, index.Count);
        foreach (byte[] digest in digests) {
            Assert.True(index.Contains(digest));
            Assert.False(index.Add(digest));
        }
        Assert.Equal(digests.Length, index.Count);
    }
}
