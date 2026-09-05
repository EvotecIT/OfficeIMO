using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Latex;
using Xunit;

namespace OfficeIMO.Latex.Tests;

public sealed class LatexDocumentIOTests {
    [Fact]
    public void ParseAndLoadExposeDirectAndStructuredLifecycleResults() {
        const string source = "\\documentclass{article}\n\\begin{document}\nBody\n\\end{document}\n";

        LatexDocument direct = LatexDocument.Parse(source);
        LatexParseResult parsed = LatexDocument.ParseResult(source);
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(source));
        LatexDocument loaded = LatexDocument.Load(stream);

        Assert.Equal(source, direct.ToLatex());
        Assert.Equal(source, loaded.ToLatex());
        Assert.IsAssignableFrom<IOfficeResult<LatexDocument>>(parsed);
        Assert.Same(parsed.Document, parsed.Value);
        Assert.Same(parsed.Document, parsed.RequireValue());
    }

    [Fact]
    public async Task StreamLifecycle_UsesCompleteArtifactsAndLeavesCallerStreamsOpen() {
        const string source = "\\documentclass{article}\n\\begin{document}\nBody\n\\end{document}\n";
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(source));
        input.Position = 5;

        LatexParseResult loaded = await LatexDocument.LoadResultAsync(input);

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
        LatexDocument document = LatexDocument.ParseResult("Body\n").Document;
        using var output = new MemoryStream(new byte[] { 1, 2, 3 });
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            document.SaveAsync(output, cancellationToken: cancellation.Token));

        Assert.Equal(new byte[] { 1, 2, 3 }, output.ToArray());
    }

    [Fact]
    public void Load_Rejects_Oversized_Seekable_Input_Before_Read_And_Restores_Position() {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("0123456789"));
        input.Position = 4;

        var options = new LatexParseOptions { MaximumInputBytes = 5 };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            LatexDocument.LoadResult(input, options));

        Assert.Contains("maximum size", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(4, input.Position);
    }

    [Fact]
    public void Load_Strips_The_Utf8_Preamble_With_Default_Encoding() {
        const string source = "\\documentclass{article}\nBody\n";
        byte[] payload = Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(source)).ToArray();
        using var input = new MemoryStream(payload);

        LatexParseResult result = LatexDocument.LoadResult(input);

        Assert.Equal(source, result.Document.ToLatex());
    }

    [Fact]
    public void FileLoadDetectsUtf16AndUtf32ByteOrderMarksWhenEncodingIsOmitted() {
        const string source = "\\documentclass{article}\nZażółć\n";
        Encoding[] encodings = {
            Encoding.Unicode,
            Encoding.BigEndianUnicode,
            new UTF32Encoding(bigEndian: false, byteOrderMark: true),
            new UTF32Encoding(bigEndian: true, byteOrderMark: true)
        };

        foreach (Encoding encoding in encodings) {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".tex");
            try {
                byte[] payload = encoding.GetPreamble().Concat(encoding.GetBytes(source)).ToArray();
                File.WriteAllBytes(path, payload);

                LatexParseResult result = LatexDocument.LoadResult(path);

                Assert.Equal(source, result.Document.ToLatex());
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }

    [Fact]
    public void Parse_Honors_PreCanceled_Token() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            LatexDocument.ParseResult("Body\n", options: null, cancellation.Token));
    }

    [Fact]
    public async Task LoadAsync_Honors_PreCanceled_Token_Without_Mutating_Stream_Position() {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("Body\n"));
        input.Position = 2;
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            LatexDocument.LoadResultAsync(input, cancellationToken: cancellation.Token));

        Assert.Equal(2, input.Position);
    }
}
