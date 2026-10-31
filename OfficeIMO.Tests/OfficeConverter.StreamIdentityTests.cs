using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Tests;

public sealed class OfficeConverterStreamIdentityTests {
    [Theory]
    [InlineData("file-stream", false)]
    [InlineData("file-stream", true)]
    [InlineData("stream-stream", false)]
    [InlineData("stream-stream", true)]
    [InlineData("stream-file", false)]
    [InlineData("stream-file", true)]
    public async Task ConversionRejectsTheSamePhysicalInputBeforeWriting(string route, bool hardLink) {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Root, "source.docx");
        byte[] original = Encoding.UTF8.GetBytes("# Preserve this input\n\nMarkdown with an explicit source format.");
        File.WriteAllBytes(input, original);
        string output = hardLink ? Path.Combine(scope.Root, "alias.docx") : input;
        if (hardLink) CreateHardLink(output, input);
        var options = new OfficeConversionOptions { SourceFormat = OfficeDocumentFormat.Markdown, FileConflictPolicy = OfficePublicationPolicy.Replace };
        // Handle-created streams deliberately have no filename: identity must come from the open handle.
        using var sourceHandle = File.OpenHandle(input, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var source = new FileStream(new SafeFileHandle(sourceHandle.DangerousGetHandle(), ownsHandle: false), FileAccess.Read);
        source.Position = 3;
        if (route == "stream-file") {
            await Assert.ThrowsAsync<ArgumentException>(() => OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Markdown, output, options));
        } else {
            using var destination = new FileStream(output, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete);
            destination.Position = 2;
            if (route == "file-stream")
                await Assert.ThrowsAsync<ArgumentException>(() => OfficeConverter.ConvertAsync(input, destination, OfficeDocumentFormat.Word, options));
            else
                await Assert.ThrowsAsync<ArgumentException>(() => OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Markdown, destination, OfficeDocumentFormat.Word, options));
            Assert.Equal(2, destination.Position);
        }
        Assert.Equal(3, source.Position);
        Assert.Equal(original, File.ReadAllBytes(input));
        Assert.Equal(original, File.ReadAllBytes(output));
    }

    [Fact]
    public async Task AnAnonymousInputCanPublishToADifferentFile() {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Root, "source.md"), output = Path.Combine(scope.Root, "output.docx");
        File.WriteAllText(input, "# Anonymous stream\n\nA distinct destination.");
        using var sourceHandle = File.OpenHandle(input);
        using var source = new FileStream(new SafeFileHandle(sourceHandle.DangerousGetHandle(), ownsHandle: false), FileAccess.Read);
        source.Position = 3;
        await OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Markdown, output);
        Assert.Equal(3, source.Position);
        using var document = WordDocument.Load(output);
        Assert.Contains(document.Paragraphs, paragraph => paragraph.Text == "Anonymous stream");
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task ImageSidecarsCannotReplaceTheInputOrCallerOwnedOutput(bool collideWithOutput, bool hardLink) {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Root, "source.docx"), output = Path.Combine(scope.Root, "output.md");
        byte[] pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        using (var document = WordDocument.Create()) {
            using var image = new MemoryStream(pixel);
            document.AddParagraph().AddImage(image, "pixel.png", OfficeLength.Pixels(10), OfficeLength.Pixels(10)).FileName = hardLink ? "sidecar.png" : collideWithOutput ? "output.md" : "source.docx";
            document.Save(input);
        }
        byte[] original = File.ReadAllBytes(input), existingOutput = Encoding.UTF8.GetBytes("Keep the caller's output");
        File.WriteAllBytes(output, existingOutput);
        if (hardLink) CreateHardLink(Path.Combine(scope.Root, "sidecar.png"), collideWithOutput ? output : input);
        using var destinationHandle = File.OpenHandle(output, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete);
        using var destination = new FileStream(new SafeFileHandle(destinationHandle.DangerousGetHandle(), ownsHandle: false), FileAccess.ReadWrite);
        destination.Position = 2;
        await Assert.ThrowsAsync<OfficeConversionException>(() => OfficeConverter.ConvertAsync(input, destination, OfficeDocumentFormat.Markdown,
            new OfficeConversionOptions {
                FileConflictPolicy = OfficePublicationPolicy.Replace,
                WordMarkdown = new WordMarkdownExportOptions { ImageExportMode = WordMarkdownImageExportMode.File, ImageDirectory = scope.Root }
            }));
        Assert.Equal(2, destination.Position);
        Assert.Equal(original, File.ReadAllBytes(input));
        Assert.Equal(existingOutput, File.ReadAllBytes(output));
    }

    private static void CreateHardLink(string alias, string source) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            if (!CreateHardLinkWindows(alias, source, IntPtr.Zero)) throw new IOException("Unable to create a test hard link.");
        } else if (Link(source, alias) != 0) throw new IOException("Unable to create a test hard link.");
    }

    [DllImport("kernel32.dll", EntryPoint = "CreateHardLinkW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateHardLinkWindows(string fileName, string existingFileName, IntPtr securityAttributes);
    [DllImport("libc", EntryPoint = "link", CharSet = CharSet.Ansi, SetLastError = true)]
    private static extern int Link(string existingPath, string newPath);

    private sealed class TestDirectory : IDisposable {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), "officeimo-stream-identity-" + Guid.NewGuid().ToString("N"));
        internal TestDirectory() => Directory.CreateDirectory(Root);
        public void Dispose() => Directory.Delete(Root, true);
    }
}
