using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.AsciiDoc;
using Xunit;

namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocDocumentIOTests {
    [Fact]
    public void ParseAndLoadExposeDirectAndStructuredLifecycleResults() {
        const string source = "= Title\n\nBody\n";

        AsciiDocDocument direct = AsciiDocDocument.Parse(source);
        AsciiDocParseResult parsed = AsciiDocDocument.ParseResult(source);
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(source));
        AsciiDocDocument loaded = AsciiDocDocument.Load(stream);

        Assert.Equal(source, direct.ToAsciiDoc());
        Assert.Equal(source, loaded.ToAsciiDoc());
        Assert.IsAssignableFrom<IOfficeResult<AsciiDocDocument>>(parsed);
        Assert.Same(parsed.Document, parsed.Value);
        Assert.Same(parsed.Document, parsed.RequireValue());
    }

    [Fact]
    public async Task StreamLifecycle_UsesCompleteArtifactsAndLeavesCallerStreamsOpen() {
        const string source = "= Title\n\nBody\n";
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(source));
        input.Position = 4;

        AsciiDocParseResult loaded = await AsciiDocDocument.LoadResultAsync(input);

        Assert.Equal(4, input.Position);
        input.ReadByte();
        Assert.Equal(source, loaded.Document.ToAsciiDoc());

        using var output = new MemoryStream(new byte[128], writable: true);
        output.Position = 19;
        await loaded.Document.SaveAsync(output);

        Assert.Equal(0, output.Position);
        Assert.Equal(loaded.Document.ToBytes(), output.ToArray());
        output.WriteByte(0);
    }

    [Fact]
    public async Task AsyncLifecycle_HonorsPreCanceledTokensWithoutMutatingStreams() {
        AsciiDocDocument document = AsciiDocDocument.ParseResult("Body\n").Document;
        using var output = new MemoryStream(new byte[] { 1, 2, 3 });
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            document.SaveAsync(output, cancellationToken: cancellation.Token));

        Assert.Equal(new byte[] { 1, 2, 3 }, output.ToArray());
    }
}
