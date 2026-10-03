using OfficeIMO.Reader;
using System.Threading;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderMaterializedCancellationTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Materialized_enumeration_observes_cancellation_between_pulls(bool pathInput, bool processor) {
        var builder = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "two-chunks", Kind = ReaderInputKind.Text, Extensions = new[] { ".rdr" },
            ReadPath = (path, _, _) => CreateChunks(path),
            ReadStream = (_, name, _, _) => CreateChunks(name)
        });
        if (processor) builder.AddProcessor(new DelegateOfficeDocumentProcessor("pass", (result, _) => result));
        OfficeDocumentReader reader = builder.Build();
        string path = Path.Combine(Path.GetTempPath(), "reader-pulls-" + Guid.NewGuid().ToString("N") + ".rdr");
        if (pathInput) File.WriteAllText(path, "input");
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("input"));
        using var cancellation = new CancellationTokenSource();
        try {
            using var iterator = (pathInput
                ? reader.EnumerateChunks(path, cancellationToken: cancellation.Token)
                : reader.EnumerateChunks(input, "input.rdr", cancellationToken: cancellation.Token)).GetEnumerator();
            Assert.True(iterator.MoveNext());
            Assert.Equal("one", iterator.Current.Text);
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => iterator.MoveNext());
        } finally { if (pathInput) File.Delete(path); }
        Assert.True(input.CanRead);
    }

    private static ReaderChunk[] CreateChunks(string name) => new[] {
        new ReaderChunk { Id = "one", Kind = ReaderInputKind.Text, Text = "one", Location = new ReaderLocation { Path = name } },
        new ReaderChunk { Id = "two", Kind = ReaderInputKind.Text, Text = "two", Location = new ReaderLocation { Path = name } }
    };
}
