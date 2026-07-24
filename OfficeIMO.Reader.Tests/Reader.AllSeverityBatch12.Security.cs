using OfficeIMO.Reader;
using OfficeIMO.Reader.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderAllSeverityBatch12SecurityTests {
    [Fact]
    public void MarkdownHandler_DefaultInputLimitRejectsOversizedPath() {
        string path = Path.Combine(Path.GetTempPath(),
            Guid.NewGuid().ToString("N") + ".md");
        try {
            using (var stream = new FileStream(path, FileMode.CreateNew,
                       FileAccess.Write, FileShare.None)) {
                stream.SetLength(
                    OfficeDocumentReaderBuilderMarkdownExtensions
                        .DefaultMaxInputBytes + 1);
            }
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddMarkdownHandler()
                .Build();

            Assert.Throws<IOException>(
                () => reader.ReadDocument(path));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void MarkdownHandler_BoundsNonSeekableStreamsBeforeParsing() {
        byte[] payload = Encoding.UTF8.GetBytes(new string('x', 1_024));
        using var stream = new NonSeekableReadStream(payload);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddMarkdownHandler()
            .Build();

        Assert.Throws<IOException>(() => reader.ReadDocument(
            stream,
            "oversized.md",
            new ReaderOptions { MaxInputBytes = 32 }));
    }

    [Fact]
    public void MarkdownDataView_OnlyMaterializesConfiguredRows() {
        var markdown = new StringBuilder(
            "```ix-dataview\n{\"columns\":[\"Value\"],\"rows\":[");
        const int totalRows = 5_000;
        for (int index = 0; index < totalRows; index++) {
            if (index > 0) markdown.Append(',');
            markdown.Append("[\"").Append(index).Append("\"]");
        }
        markdown.Append("]}\n```");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddMarkdownHandler()
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(
            Encoding.UTF8.GetBytes(markdown.ToString()),
            "table.md",
            new ReaderOptions { MaxTableRows = 2 });

        ReaderTable table = Assert.Single(
            Assert.Single(result.Chunks).Tables!);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(totalRows, table.TotalRowCount);
        Assert.True(table.Truncated);
    }

    [Fact]
    public void MarkdownDataView_CountsOnlyProjectableRows() {
        const string markdown = """
            ```ix-dataview
            {"columns":["Value"],"rows":[["one"],null,{"bad":true},["two"]]}
            ```
            """;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddMarkdownHandler()
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(
            Encoding.UTF8.GetBytes(markdown),
            "table.md",
            new ReaderOptions { MaxTableRows = 2 });

        ReaderTable table = Assert.Single(
            Assert.Single(result.Chunks).Tables!);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(2, table.TotalRowCount);
        Assert.False(table.Truncated);
    }

    [Fact]
    public void ReadFolder_RejectsEnumerationBeyondTraversalBudget() {
        string folder = Path.Combine(Path.GetTempPath(),
            "officeimo-reader-budget-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            for (int index = 0; index < 4; index++) {
                File.WriteAllText(Path.Combine(folder, index + ".md"),
                    "safe");
            }

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadFolder(
                    folder,
                    new ReaderFolderOptions {
                        Recurse = false,
                        MaxFiles = 1,
                        MaxTraversalEntries = 3
                    }).ToArray());

            Assert.Contains("MaxTraversalEntries (3)", exception.Message,
                StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, true);
        }
    }

    [Fact]
    public void ReadFolder_SkipsFileSymlinksByDefault() {
#if NET6_0_OR_GREATER
        if (OperatingSystem.IsWindows()) return;
        string root = Path.Combine(Path.GetTempPath(),
            "officeimo-reader-link-" + Guid.NewGuid().ToString("N"));
        string folder = Path.Combine(root, "input");
        string external = Path.Combine(root, "secret.md");
        Directory.CreateDirectory(folder);
        try {
            File.WriteAllText(external, "outside-secret");
            File.CreateSymbolicLink(
                Path.Combine(folder, "linked.md"), external);

            ReaderChunk[] chunks = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadFolder(
                folder,
                new ReaderFolderOptions {
                    Recurse = false,
                    Extensions = new[] { ".md" }
                }).ToArray();

            Assert.DoesNotContain(chunks, chunk =>
                (chunk.Text ?? string.Empty).Contains(
                    "outside-secret", StringComparison.Ordinal));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
#endif
    }
}
