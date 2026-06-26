using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using OfficeIMO.Markdown.Html;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderHtmlModularTests {
    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_EmitsChunks() {
        var html = "<html><body><h1>Hello HTML</h1><p>Body text.</p></body></html>";

        var chunks = DocumentReaderHtmlExtensions.ReadHtmlString(
            html: html,
            sourceName: "inline.html",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks, c =>
            c.Kind == ReaderInputKind.Html &&
            string.Equals(c.Location.Path, "inline.html", StringComparison.OrdinalIgnoreCase) &&
            ((c.Markdown ?? c.Text).Contains("Hello HTML", StringComparison.Ordinal) ||
             (c.Markdown ?? c.Text).Contains("Body text.", StringComparison.Ordinal)));
        Assert.All(chunks, c => {
            Assert.False(string.IsNullOrWhiteSpace(c.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(c.SourceHash));
            Assert.False(string.IsNullOrWhiteSpace(c.ChunkHash));
            Assert.True(c.TokenEstimate.HasValue && c.TokenEstimate.Value >= 1);
            Assert.Equal(Encoding.UTF8.GetByteCount(html), c.SourceLengthBytes);
            Assert.Null(c.SourceLastWriteUtc);
        });
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_TrimsLogicalSourceName() {
        var html = "<html><body><h1>Hello HTML</h1><p>Body text.</p></body></html>";

        var chunk = Assert.Single(DocumentReaderHtmlExtensions.ReadHtmlString(
            html: html,
            sourceName: " inline.html ",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }));

        Assert.Equal("inline.html", chunk.Location.Path);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_UsesHeadingAwareLocations() {
        var html = "<html><body><h1>Hello HTML</h1><p>Body text.</p><h2>Second</h2><p>More.</p></body></html>";

        var chunks = DocumentReaderHtmlExtensions.ReadHtmlString(
            html: html,
            sourceName: "headings.html",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks, c => !string.IsNullOrWhiteSpace(c.Location.HeadingPath));
        Assert.All(chunks, c => Assert.True(c.Location.StartLine.GetValueOrDefault() >= 1));
        Assert.Contains(chunks, c => (c.Location.HeadingPath ?? string.Empty).Contains("Hello HTML", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_CanDisableHeadingChunking() {
        var html = "<html><body><h1>Hello HTML</h1><p>Body text.</p><h2>Second</h2><p>More.</p></body></html>";

        var chunks = DocumentReaderHtmlExtensions.ReadHtmlString(
            html: html,
            sourceName: "headings-disabled.html",
            readerOptions: new ReaderOptions {
                MaxChars = 8_000,
                MarkdownChunkByHeadings = false
            }).ToList();

        Assert.NotEmpty(chunks);
        Assert.DoesNotContain(chunks, c => !string.IsNullOrWhiteSpace(c.Location.HeadingPath));
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_SplitsByMaxChars() {
        var largeHtml = "<html><body><p>" + new string('a', 2048) + "</p></body></html>";

        var chunks = DocumentReaderHtmlExtensions.ReadHtmlString(
            html: largeHtml,
            sourceName: "large.html",
            readerOptions: new ReaderOptions { MaxChars = 128 }).ToList();

        Assert.True(chunks.Count > 1);
        Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Html, c.Kind));
        Assert.Contains(chunks, c =>
            c.Warnings != null &&
            c.Warnings.Any(w => w.Contains("split due to MaxChars", StringComparison.OrdinalIgnoreCase)));
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_HeadingSplits_DoNotEmitMaxCharsWarning() {
        var html = "<html><body><h1>One</h1><p>Alpha.</p><h1>Two</h1><p>Beta.</p></body></html>";

        var chunks = DocumentReaderHtmlExtensions.ReadHtmlString(
            html: html,
            sourceName: "headings-only.html",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList();

        Assert.True(chunks.Count > 1);
        Assert.DoesNotContain(chunks, c =>
            c.Warnings?.Any(w => w.Contains("split due to MaxChars", StringComparison.OrdinalIgnoreCase)) ?? false);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_ExtractsTableProfilesFromConvertedMarkdown() {
        var html =
            "<html><body><h1>Inventory</h1><table>" +
            "<thead><tr><th>Name</th><th>Qty</th></tr></thead>" +
            "<tbody><tr><td>Paper</td><td>10</td></tr><tr><td>Ink</td><td>2</td></tr></tbody>" +
            "</table></body></html>";

        var chunk = DocumentReaderHtmlExtensions.ReadHtmlString(
            html: html,
            sourceName: "inventory.html",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).Single(c => (c.Tables?.Count ?? 0) > 0);

        Assert.Equal(ReaderInputKind.Html, chunk.Kind);
        Assert.NotNull(chunk.Tables);
        var table = Assert.Single(chunk.Tables!);
        Assert.Equal(new[] { "Name", "Qty" }, table.Columns);
        Assert.Equal(2, table.TotalRowCount);
        Assert.Equal("Paper", table.Rows[0][0]);
        Assert.Equal("2", table.Rows[1][1]);
        Assert.Equal(2, table.ColumnProfiles.Count);
        Assert.Equal(ReaderTableColumnKind.Text, table.ColumnProfiles[0].Kind);
        Assert.Equal(ReaderTableColumnKind.Numeric, table.ColumnProfiles[1].Kind);
        Assert.True(table.ColumnProfiles[1].IsNumeric);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_ExtractsLargeTableProfilesBeforeSplitting() {
        string rows = string.Concat(Enumerable.Range(1, 40).Select(index =>
            "<tr><td>Item " + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "</td><td>" + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "</td></tr>"));
        string html =
            "<html><body><h1>Inventory</h1><table>" +
            "<thead><tr><th>Name</th><th>Qty</th></tr></thead>" +
            "<tbody>" + rows + "</tbody>" +
            "</table></body></html>";

        var chunks = DocumentReaderHtmlExtensions.ReadHtmlString(
            html: html,
            sourceName: "large-table.html",
            readerOptions: new ReaderOptions { MaxChars = 128 }).ToList();

        Assert.True(chunks.Count > 1);
        var table = Assert.Single(chunks.SelectMany(chunk => chunk.Tables ?? Array.Empty<ReaderTable>()));
        Assert.Equal(new[] { "Name", "Qty" }, table.Columns);
        Assert.Equal(40, table.TotalRowCount);
        Assert.Equal("Item 1", table.Rows[0][0]);
        Assert.Equal("40", table.Rows[39][1]);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_EmitsWarningForEmptyContent() {
        var chunks = DocumentReaderHtmlExtensions.ReadHtmlString(
            html: "<html><body></body></html>",
            sourceName: "empty.html").ToList();

        Assert.Single(chunks);
        var warning = chunks[0];
        Assert.Equal("html-warning-0000", warning.Id);
        Assert.Equal(ReaderInputKind.Html, warning.Kind);
        Assert.Contains("no markdown text", warning.Text ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        Assert.False(string.IsNullOrWhiteSpace(warning.SourceId));
        Assert.False(string.IsNullOrWhiteSpace(warning.SourceHash));
        Assert.False(string.IsNullOrWhiteSpace(warning.ChunkHash));
        Assert.True(warning.TokenEstimate.HasValue && warning.TokenEstimate.Value >= 1);
        Assert.Equal(Encoding.UTF8.GetByteCount("<html><body></body></html>"), warning.SourceLengthBytes);
        Assert.Null(warning.SourceLastWriteUtc);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_PreservesConfiguredMaxInputCharacters() {
        var html = "<html><body><p>Too much content for this limit.</p></body></html>";

        var ex = Assert.Throws<ArgumentOutOfRangeException>(() => DocumentReaderHtmlExtensions.ReadHtmlString(
            html: html,
            sourceName: "limited.html",
            htmlOptions: new ReaderHtmlOptions {
                HtmlToMarkdownOptions = new HtmlToMarkdownOptions {
                    MaxInputCharacters = 12
                }
            }).ToList());

        Assert.Contains("MaxInputCharacters", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ReaderHtmlOptions_ProfilesExposeExpectedMarkdownOptions() {
        var officeProfile = ReaderHtmlOptions.CreateOfficeIMOProfile();
        Assert.NotNull(officeProfile.HtmlToMarkdownOptions);
        Assert.Null(officeProfile.HtmlToMarkdownOptions.MarkdownWriteOptions);
        Assert.Null(officeProfile.HtmlToMarkdownOptions.MaxInputCharacters);

        var portableProfile = ReaderHtmlOptions.CreatePortableProfile();
        Assert.NotNull(portableProfile.HtmlToMarkdownOptions);
        Assert.NotNull(portableProfile.HtmlToMarkdownOptions.MarkdownWriteOptions);
        Assert.Null(portableProfile.HtmlToMarkdownOptions.MaxInputCharacters);

        var untrustedProfile = ReaderHtmlOptions.CreateUntrustedHtmlProfile(64);
        Assert.NotNull(untrustedProfile.HtmlToMarkdownOptions);
        Assert.NotNull(untrustedProfile.HtmlToMarkdownOptions.MarkdownWriteOptions);
        Assert.Equal(64, untrustedProfile.HtmlToMarkdownOptions.MaxInputCharacters);

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() => ReaderHtmlOptions.CreateUntrustedHtmlProfile(0));
        Assert.Equal("maxInputCharacters", exception.ParamName);
    }

    [Fact]
    public void ReaderHtmlOptions_CloneCopiesNestedOptionsIndependently() {
        var options = ReaderHtmlOptions.CreateUntrustedHtmlProfile(64);
        options.HtmlToMarkdownOptions!.BaseUri = new Uri("https://example.com/docs/");

        var clone = options.Clone();

        Assert.NotSame(options, clone);
        Assert.NotNull(clone.HtmlToMarkdownOptions);
        Assert.NotSame(options.HtmlToMarkdownOptions, clone.HtmlToMarkdownOptions);
        Assert.Equal(options.HtmlToMarkdownOptions.BaseUri, clone.HtmlToMarkdownOptions.BaseUri);
        Assert.Equal(options.HtmlToMarkdownOptions.MaxInputCharacters, clone.HtmlToMarkdownOptions.MaxInputCharacters);
        Assert.NotNull(clone.HtmlToMarkdownOptions.MarkdownWriteOptions);
        Assert.NotSame(options.HtmlToMarkdownOptions.MarkdownWriteOptions, clone.HtmlToMarkdownOptions.MarkdownWriteOptions);

        clone.HtmlToMarkdownOptions.MaxInputCharacters = 128;

        Assert.Equal(64, options.HtmlToMarkdownOptions.MaxInputCharacters);
        Assert.Equal(128, clone.HtmlToMarkdownOptions.MaxInputCharacters);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_UntrustedProfileEnforcesMaxInputCharacters() {
        var html = "<html><body><p>Too much content for this profile limit.</p></body></html>";

        var ex = Assert.Throws<ArgumentOutOfRangeException>(() => DocumentReaderHtmlExtensions.ReadHtmlString(
            html: html,
            sourceName: "profile-limited.html",
            htmlOptions: ReaderHtmlOptions.CreateUntrustedHtmlProfile(12)).ToList());

        Assert.Contains("MaxInputCharacters", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlStream_EmitsLogicalSourceMetadata() {
        var html = "<html><body><h2>Registry HTML</h2><p>From stream.</p></body></html>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(html), writable: false);

        var chunks = DocumentReaderHtmlExtensions.ReadHtml(
            stream,
            sourceName: " metadata.html ",
            readerOptions: new ReaderOptions { MaxChars = 8_000, ComputeHashes = true }).ToList();

        Assert.NotEmpty(chunks);
        Assert.All(chunks, c => {
            Assert.Equal("metadata.html", c.Location.Path);
            Assert.False(string.IsNullOrWhiteSpace(c.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(c.SourceHash));
            Assert.False(string.IsNullOrWhiteSpace(c.ChunkHash));
            Assert.True(c.TokenEstimate.HasValue && c.TokenEstimate.Value >= 1);
            Assert.Equal(stream.Length, c.SourceLengthBytes);
            Assert.Null(c.SourceLastWriteUtc);
        });
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlFile_EmitsFileSourceMetadata() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-html-meta-" + Guid.NewGuid().ToString("N") + ".html");
        try {
            File.WriteAllText(path, "<html><body><h1>File HTML</h1><p>Body text.</p></body></html>");

            var chunks = DocumentReaderHtmlExtensions.ReadHtmlFile(
                path,
                readerOptions: new ReaderOptions { ComputeHashes = true, MaxChars = 8_000 }).ToList();

            Assert.NotEmpty(chunks);
            Assert.All(chunks, c => {
                Assert.False(string.IsNullOrWhiteSpace(c.SourceId));
                Assert.False(string.IsNullOrWhiteSpace(c.SourceHash));
                Assert.False(string.IsNullOrWhiteSpace(c.ChunkHash));
                Assert.True(c.TokenEstimate.HasValue && c.TokenEstimate.Value >= 1);
                Assert.True(c.SourceLengthBytes.HasValue && c.SourceLengthBytes.Value > 0);
                Assert.True(c.SourceLastWriteUtc.HasValue);
            });
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReaderHtml_Registration_DispatchesHtmlStream() {
        try {
            DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler(replaceExisting: true);

            var html = "<html><body><h2>Registry HTML</h2><p>From stream.</p></body></html>";
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(html), writable: false);
            var chunks = DocumentReader.Read(stream, "registry.html").ToList();

            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Html &&
                string.Equals(c.Location.Path, "registry.html", StringComparison.OrdinalIgnoreCase) &&
                ((c.Markdown ?? c.Text).Contains("Registry HTML", StringComparison.Ordinal) ||
                 (c.Markdown ?? c.Text).Contains("From stream.", StringComparison.Ordinal)));
        } finally {
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
        }
    }

    [Fact]
    public void DocumentReaderHtml_Registration_DispatchesXhtmlStream() {
        try {
            DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler(replaceExisting: true);

            var html = "<html><body><h2>Registry XHTML</h2><p>From stream.</p></body></html>";
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(html), writable: false);
            var chunks = DocumentReader.Read(stream, "registry.xhtml").ToList();

            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Html &&
                string.Equals(c.Location.Path, "registry.xhtml", StringComparison.OrdinalIgnoreCase) &&
                ((c.Markdown ?? c.Text).Contains("Registry XHTML", StringComparison.Ordinal) ||
                 (c.Markdown ?? c.Text).Contains("From stream.", StringComparison.Ordinal)));
        } finally {
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
        }
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlStream_NonSeekable_EnforcesMaxInputBytes() {
        var html = "<html><body><h2>Registry HTML</h2><p>From stream.</p></body></html>";
        using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(html));

        var ex = Assert.Throws<IOException>(() => DocumentReaderHtmlExtensions.ReadHtml(
            stream,
            sourceName: "nonseekable.html",
            readerOptions: new ReaderOptions { MaxInputBytes = 16 }).ToList());

        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlFile_EnforcesMaxInputBytes() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-html-" + Guid.NewGuid().ToString("N") + ".html");
        try {
            File.WriteAllText(path, "<html><body><p>" + new string('a', 256) + "</p></body></html>");

            var ex = Assert.Throws<IOException>(() => DocumentReaderHtmlExtensions.ReadHtmlFile(
                path,
                readerOptions: new ReaderOptions { MaxInputBytes = 32 }).ToList());

            Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReaderHtml_Registration_NonSeekableStream_EnforcesMaxInputBytes() {
        try {
            DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler(replaceExisting: true);

            var html = "<html><body><h2>Registry HTML</h2><p>From stream.</p></body></html>";
            using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(html));
            var ex = Assert.Throws<IOException>(() => DocumentReader.Read(
                stream,
                "registry.html",
                new ReaderOptions { MaxInputBytes = 16 }).ToList());

            Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
        } finally {
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
        }
    }

    [Fact]
    public void DocumentReaderHtml_Registration_PreservesConfiguredMaxInputCharacters() {
        try {
            DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler(
                htmlOptions: new ReaderHtmlOptions {
                    HtmlToMarkdownOptions = new HtmlToMarkdownOptions {
                        MaxInputCharacters = 12
                    }
                },
                replaceExisting: true);

            var html = "<html><body><p>Too much content for this limit.</p></body></html>";
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(html), writable: false);
            var ex = Assert.Throws<ArgumentOutOfRangeException>(() => DocumentReader.Read(stream, "configured.html").ToList());

            Assert.Contains("MaxInputCharacters", ex.Message, StringComparison.Ordinal);
        } finally {
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
        }
    }

    [Fact]
    public void DocumentReaderHtml_Registration_AppliesConfiguredMarkdownOptions() {
        try {
            DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler(
                htmlOptions: new ReaderHtmlOptions {
                    HtmlToMarkdownOptions = new HtmlToMarkdownOptions {
                        BaseUri = new Uri("https://example.com/docs/")
                    }
                },
                replaceExisting: true);

            var html = "<html><body><p><a href=\"guide/getting-started\">Docs</a></p></body></html>";
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(html), writable: false);
            var chunks = DocumentReader.Read(stream, "configured.html").ToList();

            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c =>
                ((c.Markdown ?? c.Text).Contains("https://example.com/docs/guide/getting-started", StringComparison.OrdinalIgnoreCase)));
        } finally {
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
        }
    }
}
