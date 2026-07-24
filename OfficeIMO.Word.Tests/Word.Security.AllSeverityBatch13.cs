using System.Diagnostics;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordAllSeverityBatch13SecurityTests {
    [Fact]
    public void MarkdownConversionProcessesEqualRunGroupsInLinearTime() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        for (int index = 0; index < 3000; index++) {
            paragraph.AddText("x");
        }

        var stopwatch = Stopwatch.StartNew();
        string markdown = document.ToMarkdown();
        stopwatch.Stop();

        Assert.Equal(3000, markdown.Count(character => character == 'x'));
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(10), $"Conversion took {stopwatch.Elapsed}.");
    }

    [Fact]
    public void MarkdownConversionCapsBlockquoteDepth() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("bounded");
        paragraph.IndentationBefore = int.MaxValue;

        string markdown = document.ToMarkdown();

        Assert.Equal(64, markdown.TakeWhile(character => character == '>' || character == ' ').Count(character => character == '>'));
        Assert.Contains("bounded", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownImageExportSanitizesHostileFileNames() {
        string imagePath = Path.Combine(AppContext.BaseDirectory, "Images", "Kulek.jpg");
        string root = Path.Combine(Path.GetTempPath(), "officeimo-md-image-" + Guid.NewGuid().ToString("N"));
        string output = Path.Combine(root, "images");
        string escaped = Path.Combine(root, "logo.jpg");
        Directory.CreateDirectory(output);
        try {
            using WordDocument source = WordDocument.Create();
            source.AddParagraph().AddImage(imagePath, 16, 16);
            source.AddParagraph().AddImage(imagePath, 16, 16);
            using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(), writable: false));
            WordImage[] images = document.Images.ToArray();
            images[0].FileName = "../a/logo.jpg";
            images[1].FileName = "../b/logo.jpg";

            string markdown = document.ToMarkdown(new WordToMarkdownOptions {
                ImageExportMode = ImageExportMode.File,
                ImageDirectory = output
            });

            Assert.False(File.Exists(escaped));
            string[] exported = Directory.GetFiles(output);
            Assert.Equal(2, exported.Length);
            Assert.Equal(2, exported.Select(Path.GetFileName).Distinct(StringComparer.OrdinalIgnoreCase).Count());
            Assert.All(exported, path => Assert.StartsWith(
                Path.GetFullPath(output) + Path.DirectorySeparatorChar,
                Path.GetFullPath(path),
                StringComparison.Ordinal));
            Assert.Contains("logo.jpg", markdown, StringComparison.Ordinal);
            Assert.Contains("logo-2.jpg", markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("..", markdown, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
    }

    [Fact]
    public void MarkdownBase64EmbeddingEnforcesImageByteLimitBeforeMaterialization() {
        string imagePath = Path.Combine(AppContext.BaseDirectory, "Images", "Kulek.jpg");
        using WordDocument source = WordDocument.Create();
        source.AddParagraph().AddImage(imagePath, 16, 16);
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(), writable: false));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            document.ToMarkdown(new WordToMarkdownOptions { MaxEmbeddedImageBytes = 1 }));

        Assert.Contains("embedded image", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void FluentRegexSearchHonorsPerMatchTimeout() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph(new string('a', 50_000) + "!");

        Assert.Throws<RegexMatchTimeoutException>(() =>
            document.AsFluent().FindRegex("^(a+)+$", TimeSpan.FromMilliseconds(1), _ => { }));
    }

    [Fact]
    public void ListTraversalRejectsHostileNumberingLevels() {
        using var stream = new MemoryStream();
        using (WordprocessingDocument package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document)) {
            MainDocumentPart main = package.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(
                    new ParagraphProperties(
                        new NumberingProperties(
                            new NumberingLevelReference { Val = int.MaxValue },
                            new NumberingId { Val = 1 })),
                    new Run(new Text("item")))));
            main.Document.Save();

            Assert.Throws<InvalidDataException>(() => WordListTraversal.Traverse(package, 8).ToList());
        }

        stream.Position = 0;
        using WordDocument document = WordDocument.Load(stream);
        Assert.Throws<InvalidDataException>(() =>
            document.ToMarkdown(new WordToMarkdownOptions { MaxListNestingDepth = 8 }));
        Assert.Throws<InvalidDataException>(() =>
            document.ExtractMarkdownChunks(
                new WordToMarkdownOptions { MaxListNestingDepth = 8 }).ToList());
    }

    [Fact]
    public void AddImageVmlRejectsNonImagePayloadWithoutAddingAnImagePart() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".png");
        File.WriteAllText(path, "not an image");
        try {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph();

            Assert.Throws<InvalidDataException>(() => paragraph.AddImageVml(path, 16, 16));
            Assert.Empty(document.Images);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void MacroParserStopsOnCyclicDirectorySectorChain() {
        byte[] compound = CreateCyclicCompoundFile();
        Type parser = typeof(WordMacro).GetNestedType("Parser", BindingFlags.NonPublic)!;
        MethodInfo method = parser.GetMethod("GetModuleNames", BindingFlags.Static | BindingFlags.NonPublic)!;
        using var stream = new MemoryStream(compound, writable: false);

        var modules = (IReadOnlyList<string>)method.Invoke(null, new object[] { stream })!;

        Assert.Empty(modules);
    }

    private static byte[] CreateCyclicCompoundFile() {
        const int sectorSize = 512;
        byte[] data = new byte[sectorSize * 3];
        BitConverter.GetBytes(0xE11AB1A1E011CFD0UL).CopyTo(data, 0);
        BitConverter.GetBytes((ushort)9).CopyTo(data, 0x1E);
        BitConverter.GetBytes(1).CopyTo(data, 0x30);
        for (int index = 0; index < 109; index++) {
            BitConverter.GetBytes(-1).CopyTo(data, 0x4C + index * 4);
        }
        BitConverter.GetBytes(0).CopyTo(data, 0x4C);
        int fatOffset = sectorSize;
        BitConverter.GetBytes(unchecked((int)0xFFFFFFFD)).CopyTo(data, fatOffset);
        BitConverter.GetBytes(1).CopyTo(data, fatOffset + 4);
        return data;
    }
}
