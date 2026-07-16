using OfficeIMO.Epub;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderEpubModularTests {
    [Fact]
    public void EpubReader_RawHtmlCapPreservesTextlessStructuredChapter() {
        string epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-structured-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildTextlessStructuredEpub(epubPath);

            EpubDocument document = EpubDocument.Load(epubPath, new EpubReadOptions {
                IncludeRawHtml = true,
                MaxTotalRawHtmlBytes = 1
            });

            EpubChapter chapter = Assert.Single(document.Chapters);
            Assert.True(chapter.HasStructuredContent);
            Assert.Equal(string.Empty, chapter.Text);
            Assert.Null(chapter.Html);
            Assert.Contains(document.Warnings, warning => warning.Contains("MaxTotalRawHtmlBytes", StringComparison.Ordinal));
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_KeepsHtmlNoMarkdownWarningOutOfChapterContent() {
        string epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-structured-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildTextlessStructuredEpub(epubPath);

            ReaderChunk chapter = Assert.Single(EpubReaderAdapter.Read(epubPath));

            Assert.Equal(string.Empty, chapter.Text);
            Assert.Contains("## Canvas Chapter", chapter.Markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("no markdown text", chapter.Markdown, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(chapter.Warnings!, warning => warning.Contains("no markdown text", StringComparison.OrdinalIgnoreCase));
            Assert.Equal("chapter", chapter.Location.SourceBlockKind);
            Assert.EndsWith("::OEBPS/chapter.xhtml", chapter.Location.Path, StringComparison.Ordinal);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    private static void BuildTextlessStructuredEpub(string epubPath) {
        using var file = new FileStream(epubPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create, leaveOpen: false);
        WriteTextEntry(
            archive,
            "META-INF/container.xml",
            "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\"><rootfiles>" +
            "<rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/>" +
            "</rootfiles></container>");
        WriteTextEntry(
            archive,
            "OEBPS/content.opf",
            "<package version=\"3.0\" xmlns=\"http://www.idpf.org/2007/opf\"><manifest>" +
            "<item id=\"chapter\" href=\"chapter.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "</manifest><spine><itemref idref=\"chapter\"/></spine></package>");
        WriteTextEntry(
            archive,
            "OEBPS/chapter.xhtml",
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Canvas Chapter</title></head>" +
            "<body><form aria-label=\"Interactive form\"></form></body></html>");
    }
}
