using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

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
            c.Kind == ReaderInputKind.Unknown &&
            string.Equals(c.Location.Path, "inline.html", StringComparison.OrdinalIgnoreCase) &&
            ((c.Markdown ?? c.Text).Contains("Hello HTML", StringComparison.Ordinal) ||
             (c.Markdown ?? c.Text).Contains("Body text.", StringComparison.Ordinal)));
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_SplitsByMaxChars() {
        var largeHtml = "<html><body><p>" + new string('a', 2048) + "</p></body></html>";

        var chunks = DocumentReaderHtmlExtensions.ReadHtmlString(
            html: largeHtml,
            sourceName: "large.html",
            readerOptions: new ReaderOptions { MaxChars = 128 }).ToList();

        Assert.True(chunks.Count > 1);
        Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Unknown, c.Kind));
        Assert.Contains(chunks, c =>
            c.Warnings != null &&
            c.Warnings.Any(w => w.Contains("split due to MaxChars", StringComparison.OrdinalIgnoreCase)));
    }

    [Fact]
    public void DocumentReaderHtml_ReadHtmlString_EmitsWarningForEmptyContent() {
        var chunks = DocumentReaderHtmlExtensions.ReadHtmlString(
            html: "<html><body></body></html>",
            sourceName: "empty.html").ToList();

        Assert.Single(chunks);
        var warning = chunks[0];
        Assert.Equal("html-warning-0000", warning.Id);
        Assert.Equal(ReaderInputKind.Unknown, warning.Kind);
        Assert.Contains("no markdown text", warning.Text ?? string.Empty, StringComparison.OrdinalIgnoreCase);
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
                c.Kind == ReaderInputKind.Unknown &&
                string.Equals(c.Location.Path, "registry.html", StringComparison.OrdinalIgnoreCase) &&
                ((c.Markdown ?? c.Text).Contains("Registry HTML", StringComparison.Ordinal) ||
                 (c.Markdown ?? c.Text).Contains("From stream.", StringComparison.Ordinal)));
        } finally {
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
        }
    }
}
