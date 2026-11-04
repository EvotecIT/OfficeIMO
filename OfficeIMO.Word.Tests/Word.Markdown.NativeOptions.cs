using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownNativeOptionsTests {
    [Fact]
    public async Task NativeRemoteImageOptInDownloadsOnceAndLeavesCallerOptionsReusable() {
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(20));
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        int port = ((IPEndPoint)listener.LocalEndpoint).Port;
        byte[] png = await File.ReadAllBytesAsync(Path.Combine(AppContext.BaseDirectory, "NativeOptionAssets", "OfficeIMO.png"), deadline.Token);
        Task server = ServeImageAsync(listener, png, deadline.Token);
        try {
            string uri = $"http://127.0.0.1:{port}/image.png";
            using MarkdownDocument source = MarkdownDocument.Parse($"![first]({uri})\n\n![second]({uri})");
            var options = new MarkdownWordExportOptions { AllowRemoteImages = true, MaximumRemoteImageBytes = png.Length,
                RemoteImageLoadOptions = new OfficeRemoteImageLoadOptions { AllowPrivateNetworkAddresses = true } };
            OfficeConversionResult<WordDocument> converted = await source.ToWordAsync(options, deadline.Token);
            using WordDocument word = converted.Value;
            await server.WaitAsync(deadline.Token);
            Assert.True(word.Images.Count == 2, string.Join("\n", converted.Report.Diagnostics.Select(finding => finding.Message)));
            Assert.Null(options.RemoteImageResolver);
            Assert.True(source.Blocks.Count > 0);
        } finally { listener.Stop(); try { await server; } catch (OperationCanceledException) { } catch (SocketException) { } }
    }

    [Fact]
    public async Task RemoteImageByteRejectionRetainsLossAndDoesNotReplaceTheDestination() {
        using MarkdownDocument source = MarkdownDocument.Parse("![image](https://example.invalid/image.png)");
        using var output = new MemoryStream();
        var options = new MarkdownWordExportOptions { AllowRemoteImages = true, MaximumRemoteImageBytes = 0, FailOnLoss = true };
        OfficeConversionException error = await Assert.ThrowsAsync<OfficeConversionException>(() => source.SaveAsWordAsync(output, options));
        Assert.True(error.Report.HasLoss);
        Assert.Equal(0, output.Length);
        Assert.True(output.CanWrite);
    }

    [Fact]
    public async Task SharedThemeRetainsAuthoredFontsAndColorsInHtmlAndWordArtifacts() {
        OfficeTheme native = OfficeTheme.Modern.WithFonts(OfficeTheme.Modern.Fonts.With(
            heading: OfficeTheme.Modern.Fonts.Heading.With(family: "Georgia"), body: OfficeTheme.Modern.Fonts.Body.With(family: "Courier New")))
            .WithPalette(OfficeTheme.Modern.Palette.With(accent1: OfficeColor.Parse("#13579B")));
        MarkdownVisualTheme mapped = MarkdownVisualTheme.FromOfficeTheme(native);
        using MarkdownDocument source = MarkdownDocument.Parse("# Heading\n\nBody text.");
        string html = source.ToHtml(new MarkdownHtmlExportOptions { Theme = mapped, Kind = MarkdownHtmlKind.Document });
        Assert.Contains("Courier New", html, StringComparison.Ordinal);
        Assert.Contains("13579B", html, StringComparison.OrdinalIgnoreCase);
        var converted = await source.ToWordAsync(new MarkdownWordExportOptions { Theme = mapped });
        using WordDocument word = converted.Value;
        using var bytes = new MemoryStream(word.ToBytes());
        using var package = new ZipArchive(bytes, ZipArchiveMode.Read);
        using var styles = new StreamReader(package.GetEntry("word/styles.xml")!.Open());
        Assert.Contains("Georgia", await styles.ReadToEndAsync(), StringComparison.Ordinal);
        using var theme = new StreamReader(package.GetEntry("word/theme/theme1.xml")!.Open());
        string themeXml = await theme.ReadToEndAsync();
        Assert.Contains("Courier New", themeXml, StringComparison.Ordinal);
        Assert.Contains("13579B", themeXml, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(native.TypeScale.Heading1, mapped.Clone().SourceTheme!.TypeScale.Heading1);
    }

    private static async Task ServeImageAsync(TcpListener listener, byte[] image, CancellationToken token) {
        using TcpClient client = await listener.AcceptTcpClientAsync(token);
        using NetworkStream stream = client.GetStream();
        var bytes = new List<byte>();
        var buffer = new byte[1024];
        while (!Encoding.ASCII.GetString(bytes.ToArray()).Contains("\r\n\r\n", StringComparison.Ordinal)) {
            int read = await stream.ReadAsync(buffer, token);
            if (read == 0) throw new EndOfStreamException();
            bytes.AddRange(buffer.Take(read));
            if (bytes.Count > 16384) throw new InvalidDataException("Unexpected loopback request size.");
        }
        byte[] header = Encoding.ASCII.GetBytes($"HTTP/1.1 200 OK\r\nContent-Type: image/png\r\nContent-Length: {image.Length}\r\nConnection: close\r\n\r\n");
        await stream.WriteAsync(header, token);
        await stream.WriteAsync(image, token);
    }
}
