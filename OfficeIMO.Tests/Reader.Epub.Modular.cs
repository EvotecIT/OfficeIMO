using OfficeIMO.Epub;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderEpubModularTests {
    [Fact]
    public void EpubReader_UsesOpfSpineOrderAndMetadata() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath);

            var document = EpubReader.Read(epubPath, new EpubReadOptions {
                PreferSpineOrder = true,
                IncludeNonLinearSpineItems = true
            });

            Assert.NotNull(document);
            Assert.Equal("Demo Book", document.Title);
            Assert.Equal("author-1", document.Identifier);
            Assert.Equal("en", document.Language);
            Assert.Equal("OfficeIMO Team", document.Creator);
            Assert.Equal("OEBPS/content.opf", document.OpfPath);
            Assert.Equal(2, document.Chapters.Count);

            var first = document.Chapters[0];
            var second = document.Chapters[1];

            Assert.Equal("OEBPS/chapter2.xhtml", first.Path);
            Assert.Equal("Second", first.Title);
            Assert.Equal(1, first.SpineIndex);
            Assert.True(first.IsLinear);
            Assert.Contains("Second chapter text.", first.Text, StringComparison.Ordinal);

            Assert.Equal("OEBPS/chapter1.xhtml", second.Path);
            Assert.Equal("First", second.Title);
            Assert.Equal(2, second.SpineIndex);
            Assert.True(second.IsLinear);
            Assert.Contains("First chapter text.", second.Text, StringComparison.Ordinal);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_EmitsWarningsAndVirtualChapterPaths() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithMalformedChapter(epubPath);

            var chunks = DocumentReaderEpubExtensions.ReadEpub(
                epubPath,
                readerOptions: new ReaderOptions { MaxChars = 64 },
                epubOptions: new EpubReadOptions { PreferSpineOrder = true, FallbackToHtmlScan = true }).ToList();

            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Unknown &&
                (c.Warnings?.Any(w => w.Contains("not valid XML", StringComparison.OrdinalIgnoreCase)) ?? false));

            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Unknown &&
                (c.Location.Path?.Contains("::OEBPS/good.xhtml", StringComparison.OrdinalIgnoreCase) ?? false) &&
                (c.Text?.Contains("Good chapter body text.", StringComparison.Ordinal) ?? false));
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_ReadsFromNonSeekableStream() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath);
            var bytes = File.ReadAllBytes(epubPath);

            using var stream = new NonSeekableReadStream(bytes);
            var chunks = DocumentReaderEpubExtensions.ReadEpub(
                stream,
                sourceName: "nonseekable.epub",
                readerOptions: new ReaderOptions { MaxChars = 4_000 },
                epubOptions: new EpubReadOptions { PreferSpineOrder = true }).ToList();

            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Unknown &&
                (c.Location.Path?.Contains("nonseekable.epub::OEBPS/chapter2.xhtml", StringComparison.OrdinalIgnoreCase) ?? false) &&
                (c.Text?.Contains("Second chapter text.", StringComparison.Ordinal) ?? false));
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    private static void BuildEpubWithSpine(string epubPath) {
        using var fs = new FileStream(epubPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(fs, ZipArchiveMode.Create, leaveOpen: false);

        WriteTextEntry(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
        WriteTextEntry(archive, "META-INF/container.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\">" +
            "<rootfiles><rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles>" +
            "</container>");

        WriteTextEntry(archive, "OEBPS/content.opf",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<package version=\"3.0\" unique-identifier=\"bookid\" xmlns=\"http://www.idpf.org/2007/opf\">" +
            "<metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\">" +
            "<dc:title>Demo Book</dc:title><dc:creator>OfficeIMO Team</dc:creator><dc:language>en</dc:language><dc:identifier id=\"bookid\">author-1</dc:identifier>" +
            "</metadata>" +
            "<manifest>" +
            "<item id=\"nav\" href=\"nav.xhtml\" media-type=\"application/xhtml+xml\" properties=\"nav\"/>" +
            "<item id=\"ncx\" href=\"toc.ncx\" media-type=\"application/x-dtbncx+xml\"/>" +
            "<item id=\"ch1\" href=\"chapter1.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "<item id=\"ch2\" href=\"chapter2.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "</manifest>" +
            "<spine toc=\"ncx\"><itemref idref=\"ch2\"/><itemref idref=\"ch1\"/></spine>" +
            "</package>");

        WriteTextEntry(archive, "OEBPS/nav.xhtml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><nav epub:type=\"toc\" xmlns:epub=\"http://www.idpf.org/2007/ops\"><ol>" +
            "<li><a href=\"chapter2.xhtml\">Second</a></li>" +
            "<li><a href=\"chapter1.xhtml\">First</a></li>" +
            "</ol></nav></body></html>");

        WriteTextEntry(archive, "OEBPS/toc.ncx",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<ncx xmlns=\"http://www.daisy.org/z3986/2005/ncx/\" version=\"2005-1\"><navMap>" +
            "<navPoint id=\"n1\"><navLabel><text>Second</text></navLabel><content src=\"chapter2.xhtml\"/></navPoint>" +
            "<navPoint id=\"n2\"><navLabel><text>First</text></navLabel><content src=\"chapter1.xhtml\"/></navPoint>" +
            "</navMap></ncx>");

        WriteTextEntry(archive, "OEBPS/chapter1.xhtml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Local One</title></head><body><h1>One</h1><p>First chapter text.</p></body></html>");

        WriteTextEntry(archive, "OEBPS/chapter2.xhtml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Local Two</title></head><body><h1>Two</h1><p>Second chapter text.</p></body></html>");
    }

    private static void BuildEpubWithMalformedChapter(string epubPath) {
        using var fs = new FileStream(epubPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(fs, ZipArchiveMode.Create, leaveOpen: false);

        WriteTextEntry(archive, "META-INF/container.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\">" +
            "<rootfiles><rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles>" +
            "</container>");

        WriteTextEntry(archive, "OEBPS/content.opf",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<package version=\"3.0\" xmlns=\"http://www.idpf.org/2007/opf\">" +
            "<manifest>" +
            "<item id=\"good\" href=\"good.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "<item id=\"bad\" href=\"bad.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "</manifest>" +
            "<spine><itemref idref=\"bad\"/><itemref idref=\"good\"/></spine>" +
            "</package>");

        WriteTextEntry(archive, "OEBPS/bad.xhtml", "<html><body><p>broken");
        WriteTextEntry(archive, "OEBPS/good.xhtml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><p>Good chapter body text.</p></body></html>");
    }

    private static void WriteTextEntry(ZipArchive archive, string path, string content, CompressionLevel compressionLevel = CompressionLevel.Optimal) {
        var entry = archive.CreateEntry(path, compressionLevel);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), 4096, leaveOpen: false);
        writer.Write(content);
    }

    private sealed class NonSeekableReadStream : Stream {
        private readonly Stream _inner;

        public NonSeekableReadStream(byte[] bytes) {
            _inner = new MemoryStream(bytes, writable: false);
        }

        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            if (disposing) {
                _inner.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
