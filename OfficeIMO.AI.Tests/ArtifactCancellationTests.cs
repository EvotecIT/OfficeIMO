using System.Collections;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class ArtifactCancellationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CancellationStopsExportBeforeAnotherRowOrArtifact(bool duringRows) {
        using var cancellation = new CancellationTokenSource();
        var rows = new CancellingRows(duringRows ? cancellation.Cancel : () => { });
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, new OfficeDocumentReadResult {
            Blocks = new[] { new OfficeDocumentBlock { Text = "42" } }
        });
        var result = new OfficeAiResult {
            RequestId = "cancellation", SourceHash = document.SourceHash, SnapshotHash = document.SnapshotHash,
            Operation = OfficeAiOperation.Parse, Status = OfficeAiResultStatus.Completed,
            Profile = new() { Id = "export", Provider = "fixture", Model = "fixture", IsLocal = true },
            Tables = new[] { new OfficeAiTable(new ReaderTable { Columns = new[] { "Count" }, Rows = rows },
                new[] { new OfficeAiCitation("e1", null, "42", true) }) }
        };
        string output = Path.Combine(Path.GetTempPath(), "officeimo-ai-cancel-" + Guid.NewGuid().ToString("N"));
        if (!duringRows) cancellation.Cancel();
        try {
            var error = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ArtifactWriter.SaveAsync(output, document, result, cancellation.Token));
            Assert.Equal(cancellation.Token, error.CancellationToken);
            Assert.False(File.Exists(Path.Combine(output, "Table-1.csv")));
            Assert.False(File.Exists(Path.Combine(output, "extraction.xlsx")));
            if (duringRows) Assert.Equal(1, rows.IndexedReads);
            else Assert.False(Directory.Exists(output));
        } finally { if (Directory.Exists(output)) Directory.Delete(output, recursive: true); }
    }

    // Enumerating source rows is harmless; cancellation arrives at the first indexed row write/readback.
    private sealed class CancellingRows(Action cancel) : IReadOnlyList<IReadOnlyList<string>> {
        public int IndexedReads;
        public int Count => 3;
        public IReadOnlyList<string> this[int index] { get { IndexedReads++; cancel(); return new[] { "42" }; } }
        public IEnumerator<IReadOnlyList<string>> GetEnumerator() {
            for (int i = 0; i < Count; i++) yield return new[] { "42" };
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
