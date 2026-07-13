using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Latex;
using Xunit;

namespace OfficeIMO.Latex.Tests;

public sealed class LatexDocumentIOTests {
    [Fact]
    public async Task StreamLifecycle_UsesCompleteArtifactsAndLeavesCallerStreamsOpen() {
        const string source = "\\documentclass{article}\n\\begin{document}\nBody\n\\end{document}\n";
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(source));
        input.Position = 5;

        LatexParseResult loaded = await LatexDocument.LoadAsync(input);

        Assert.Equal(5, input.Position);
        input.ReadByte();
        Assert.Equal(source, loaded.Document.ToLatex());

        using var output = new MemoryStream(new byte[128], writable: true);
        output.Position = 23;
        await loaded.Document.SaveAsync(output);

        Assert.Equal(0, output.Position);
        Assert.Equal(loaded.Document.ToBytes(), output.ToArray());
        output.WriteByte(0);
    }

    [Fact]
    public async Task AsyncLifecycle_HonorsPreCanceledTokensWithoutMutatingStreams() {
        LatexDocument document = LatexDocument.Parse("Body\n").Document;
        using var output = new MemoryStream(new byte[] { 1, 2, 3 });
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            document.SaveAsync(output, cancellationToken: cancellation.Token));

        Assert.Equal(new byte[] { 1, 2, 3 }, output.ToArray());
    }
}
