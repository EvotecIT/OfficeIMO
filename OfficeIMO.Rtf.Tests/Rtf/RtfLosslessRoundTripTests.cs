using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using OfficeIMO.Rtf.Syntax;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfLosslessRoundTripTests {
    [Fact]
    public void SyntaxTree_ToRtf_Preserves_Unknown_Destinations_Control_Spacing_And_Binary() {
        const string rtf = @"{\rtf1\ansi{\*\unknown\foo-12 bar \'80}{\object\objdata 0102}\pard Before \bin4 a{b} after\par}";

        RtfSyntaxTree tree = RtfSyntaxTree.Parse(rtf);
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.DoesNotContain(tree.Diagnostics, diagnostic => diagnostic.Severity == RtfDiagnosticSeverity.Error);
        Assert.Equal(rtf, tree.ToRtf());
        Assert.Equal(rtf, result.ToRtfLossless());
    }

    [Fact]
    public void Load_And_SaveLossless_Preserve_Raw_Binary_File_Bytes() {
        string inputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".rtf");
        string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".rtf");

        byte[] bytes = new byte[] {
            123, 92, 114, 116, 102, 49, 92, 97, 110, 115, 105, 92, 98, 105, 110, 49, 32, 0x80, 125
        };

        try {
            File.WriteAllBytes(inputPath, bytes);

            RtfReadResult result = RtfDocument.Load(inputPath);
            result.SaveLossless(outputPath);

            Assert.Equal(bytes, File.ReadAllBytes(outputPath));
            Assert.Equal(@"{\rtf1\ansi\bin1 " + (char)0x80 + "}", result.ToRtfLossless());
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void Load_And_ToBytesLossless_Preserve_Raw_Binary_Bytes() {
        byte[] bytes = new byte[] {
            123, 92, 114, 116, 102, 49, 92, 97, 110, 115, 105, 92, 98, 105, 110, 49, 32, 0x80, 125
        };

        RtfReadResult result = RtfDocument.Load(bytes);

        Assert.Equal(bytes, result.ToBytesLossless());
        Assert.Equal(@"{\rtf1\ansi\bin1 " + (char)0x80 + "}", result.ToRtfLossless());
    }

    [Fact]
    public async Task LoadAsync_And_SaveLosslessAsync_Preserve_Raw_Binary_File_Bytes() {
        string inputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".rtf");
        string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".rtf");

        byte[] bytes = new byte[] {
            123, 92, 114, 116, 102, 49, 92, 97, 110, 115, 105, 92, 98, 105, 110, 49, 32, 0x80, 125
        };

        try {
            File.WriteAllBytes(inputPath, bytes);

            RtfReadResult result = await RtfDocument.LoadAsync(inputPath);
            await result.SaveLosslessAsync(outputPath);

            Assert.Equal(bytes, File.ReadAllBytes(outputPath));
            Assert.Equal(bytes, result.ToBytesLossless());
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public async Task LoadAsync_RestoresInputAndSaveLosslessAsync_OverwritesAndRewindsOutput() {
        byte[] bytes = new byte[] {
            123, 92, 114, 116, 102, 49, 92, 97, 110, 115, 105, 92, 98, 105, 110, 49, 32, 0x80, 125
        };

        using var input = new MemoryStream(bytes);
        RtfReadResult result = await RtfDocument.LoadAsync(input);

        using var output = new MemoryStream();
        output.WriteByte(0x2A);
        await result.SaveLosslessAsync(output);
        byte[] saved = output.ToArray();

        Assert.Equal(0, input.Position);
        Assert.Equal(0, output.Position);
        Assert.Equal(bytes, saved);
    }

    [Fact]
    public async Task LosslessEditor_SaveLosslessAsync_PreservesSyntaxAndOverwritesAndRewindsOutput() {
        const string rtf = @"{\rtf1\ansi{\*\unknown Keep}{\pict\pngblip\bin3 abc}\pard Existing\par}";
        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.AppendParagraph("Next");

        using var output = new MemoryStream();
        output.WriteByte(0x2A);
        await editor.SaveLosslessAsync(output);
        byte[] saved = output.ToArray();

        Assert.Equal(0, output.Position);
        Assert.Equal(editor.ToBytesLossless(), saved);
        Assert.Contains(@"{\*\unknown Keep}", editor.ToRtf(), StringComparison.Ordinal);
        Assert.Contains(@"{\pict\pngblip\bin3 abc}", editor.ToRtf(), StringComparison.Ordinal);

        RtfReadResult edited = RtfDocument.Read(editor.ToRtf());
        Assert.Equal(editor.ToRtf(), edited.ToRtfLossless());
    }

    [Fact]
    public async Task RtfDocument_SaveAsync_Uses_Semantic_Utf8_Output_Not_Lossless_Raw_Bytes() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Semantic ż");

        using var output = new MemoryStream();
        output.WriteByte(0x2A);

        await document.SaveAsync(output, new RtfWriteOptions { IncludeGenerator = false });
        byte[] saved = output.ToArray();

        Assert.Equal(0, output.Position);
        Assert.Equal(document.ToRtf(new RtfWriteOptions { IncludeGenerator = false }), Encoding.UTF8.GetString(saved));
    }

    [Fact]
    public async Task RtfDocument_LoadAsync_Honors_Cancellation() {
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        using var input = new MemoryStream(new byte[] { 123, 125 });
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            RtfDocument.LoadAsync(input, cancellationToken: cts.Token));
    }
}
