using OfficeIMO.Epub;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubPackageContractTests {
    [Fact]
    public void Load_ExposesIdentityLayoutEncryptionAndStructuredDiagnostics() {
        byte[] package = BuildPackage(includeUnsafeEntry: false);

        EpubDocument document = EpubDocument.Load(
            new MemoryStream(package, writable: false),
            new EpubReadOptions {
                IncludeRawHtml = true,
                IncludeResourceData = true
            });

        Assert.Equal("3.0", document.PackageVersion);
        Assert.Equal("primary-id", document.UniqueIdentifierId);
        Assert.Equal("urn:book:primary", document.Identifier);
        Assert.Equal(EpubRenditionLayout.PrePaginated, document.RenditionLayout);
        Assert.True(document.IsFixedLayout);

        EpubChapter chapter = Assert.Single(document.Chapters);
        Assert.Equal(EpubRenditionLayout.Reflowable, chapter.RenditionLayout);
        Assert.False(chapter.IsFixedLayout);

        Assert.True(document.HasEncryptedResources);
        Assert.True(document.RequiresDecryption);
        Assert.Equal(2, document.Encryption.Count);
        EpubEncryptionInfo fontEncryption = Assert.Single(
            document.Encryption,
            item => item.Path == "EPUB/fonts/book.otf");
        Assert.Equal(EpubEncryptionKind.IdpfFontObfuscation, fontEncryption.Kind);
        Assert.True(fontEncryption.IsFontObfuscation);
        Assert.False(fontEncryption.RequiresDecryption);
        EpubEncryptionInfo protectedEncryption = Assert.Single(
            document.Encryption,
            item => item.Path == "EPUB/protected.bin");
        Assert.Equal(EpubEncryptionKind.Encryption, protectedEncryption.Kind);
        Assert.True(protectedEncryption.RequiresDecryption);

        EpubResource font = Assert.Single(document.Resources, item => item.Id == "font");
        Assert.NotNull(font.Data);
        Assert.Same(fontEncryption, font.Encryption);
        EpubResource protectedResource = Assert.Single(document.Resources, item => item.Id == "protected");
        Assert.Null(protectedResource.Data);
        Assert.Same(protectedEncryption, protectedResource.Encryption);

        Assert.Contains(document.Diagnostics, item =>
            item.Code == "epub.encryption.font-obfuscation" &&
            item.Severity == EpubDiagnosticSeverity.Info &&
            item.Path == "EPUB/fonts/book.otf");
        Assert.Contains(document.Diagnostics, item =>
            item.Code == "epub.encryption.unsupported" &&
            item.Severity == EpubDiagnosticSeverity.Warning &&
            item.Path == "EPUB/protected.bin");
        Assert.Contains(document.Diagnostics, item => item.Code == "epub.layout.fixed");
        Assert.DoesNotContain(document.Warnings, warning => warning.Contains("font obfuscation", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(document.Warnings, warning => warning.Contains("unsupported encryption", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task LoadAndLoadAsync_EnforcePackageSizeWithFatalDiagnostics() {
        byte[] package = BuildPackage(includeUnsafeEntry: false);
        var options = new EpubReadOptions { MaxPackageBytes = package.Length - 1L };

        EpubReadException syncException = Assert.Throws<EpubReadException>(() =>
            EpubDocument.Load(new MemoryStream(package, writable: false), options));
        EpubDiagnostic syncDiagnostic = Assert.Single(syncException.Diagnostics);
        Assert.Equal("epub.package.size-limit", syncDiagnostic.Code);
        Assert.Equal(EpubDiagnosticSeverity.Error, syncDiagnostic.Severity);

        EpubReadException asyncException = await Assert.ThrowsAsync<EpubReadException>(() =>
            EpubDocument.LoadAsync(new MemoryStream(package, writable: false), options));
        Assert.Equal("epub.package.size-limit", Assert.Single(asyncException.Diagnostics).Code);
    }

    [Fact]
    public void Load_EnforcesArchiveEntryAndDeclaredUncompressedLimits() {
        byte[] package = BuildPackage(includeUnsafeEntry: false);

        EpubReadException entryException = Assert.Throws<EpubReadException>(() =>
            EpubDocument.Load(
                new MemoryStream(package, writable: false),
                new EpubReadOptions { MaxArchiveEntries = 2 }));
        Assert.Equal("epub.archive.entry-count-limit", Assert.Single(entryException.Diagnostics).Code);

        EpubReadException sizeException = Assert.Throws<EpubReadException>(() =>
            EpubDocument.Load(
                new MemoryStream(package, writable: false),
                new EpubReadOptions { MaxTotalUncompressedBytes = 32 }));
        Assert.Equal("epub.archive.total-size-limit", Assert.Single(sizeException.Diagnostics).Code);
    }

    [Fact]
    public void Load_IgnoresUnsafeArchivePathsAndReportsStableDiagnostic() {
        byte[] package = BuildPackage(includeUnsafeEntry: true);

        EpubDocument document = EpubDocument.Load(new MemoryStream(package, writable: false));

        EpubDiagnostic[] diagnostics = document.Diagnostics
            .Where(item => item.Code == "epub.archive.unsafe-path")
            .ToArray();
        Assert.Equal(2, diagnostics.Length);
        Assert.Contains(diagnostics, item => item.Path == "../outside.xhtml");
        Assert.Contains(diagnostics, item => item.Path == "EPUB/../also-outside.xhtml");
        Assert.DoesNotContain(diagnostics, item => item.Path == "EPUB/");
        Assert.DoesNotContain(document.Chapters, chapter => chapter.Path.Contains("outside", StringComparison.Ordinal));
    }

    [Fact]
    public void Load_PreservesCaseDistinctManifestIdsAndArchivePaths() {
        byte[] package = BuildCaseSensitivePackage();

        EpubDocument document = EpubDocument.Load(new MemoryStream(package, writable: false));

        Assert.Collection(
            document.Chapters,
            chapter => {
                Assert.Equal("EPUB/Chapter.xhtml", chapter.Path);
                Assert.Equal("Chapter", chapter.ManifestId);
                Assert.Contains("Upper body", chapter.Text, StringComparison.Ordinal);
            },
            chapter => {
                Assert.Equal("EPUB/chapter.xhtml", chapter.Path);
                Assert.Equal("chapter", chapter.ManifestId);
                Assert.Contains("Lower body", chapter.Text, StringComparison.Ordinal);
            });
    }

    [Fact]
    public void Load_ParsesEpub2NcxIdentityAndLegacyFixedLayout() {
        byte[] package = BuildEpub2Package();

        EpubDocument document = EpubDocument.Load(new MemoryStream(package, writable: false));

        Assert.Equal("2.0", document.PackageVersion);
        Assert.Equal("book-id", document.UniqueIdentifierId);
        Assert.Equal("urn:epub2:book", document.Identifier);
        Assert.Equal(EpubRenditionLayout.PrePaginated, document.RenditionLayout);
        Assert.True(document.IsFixedLayout);
        EpubChapter chapter = Assert.Single(document.Chapters);
        Assert.Equal("NCX Chapter", chapter.Title);
        Assert.True(chapter.IsFixedLayout);
    }

    [Fact]
    public void Load_InvalidZipThrowsStructuredFatalDiagnostic() {
        EpubReadException exception = Assert.Throws<EpubReadException>(() =>
            EpubDocument.Load(new MemoryStream(new byte[] { 1, 2, 3, 4 }, writable: false)));

        EpubDiagnostic diagnostic = Assert.Single(exception.Diagnostics);
        Assert.Equal("epub.archive.invalid", diagnostic.Code);
        Assert.Equal(EpubDiagnosticSeverity.Error, diagnostic.Severity);
    }

    private static byte[] BuildPackage(bool includeUnsafeEntry) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteTextEntry(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
            WriteTextEntry(
                archive,
                "META-INF/container.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\">" +
                "<rootfiles><rootfile full-path=\"EPUB/package.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles>" +
                "</container>");
            WriteTextEntry(
                archive,
                "META-INF/encryption.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<encryption xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\" xmlns:enc=\"http://www.w3.org/2001/04/xmlenc#\">" +
                "<enc:EncryptedData><enc:EncryptionMethod Algorithm=\"http://www.idpf.org/2008/embedding\"/>" +
                "<enc:CipherData><enc:CipherReference URI=\"EPUB/fonts/book.otf\"/></enc:CipherData></enc:EncryptedData>" +
                "<enc:EncryptedData><enc:EncryptionMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#aes256-cbc\"/>" +
                "<enc:CipherData><enc:CipherReference URI=\"EPUB/protected.bin\"/></enc:CipherData></enc:EncryptedData>" +
                "</encryption>");
            WriteTextEntry(
                archive,
                "EPUB/package.opf",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<package version=\"3.0\" unique-identifier=\"primary-id\" xmlns=\"http://www.idpf.org/2007/opf\">" +
                "<metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\">" +
                "<dc:identifier id=\"secondary-id\">urn:book:secondary</dc:identifier>" +
                "<dc:identifier id=\"primary-id\">urn:book:primary</dc:identifier>" +
                "<dc:title>Package contracts</dc:title><dc:language>en</dc:language>" +
                "<meta property=\"rendition:layout\">pre-paginated</meta>" +
                "</metadata>" +
                "<manifest>" +
                "<item id=\"chapter\" href=\"chapter.xhtml\" media-type=\"application/xhtml+xml\"/>" +
                "<item id=\"font\" href=\"fonts/book.otf\" media-type=\"font/otf\"/>" +
                "<item id=\"protected\" href=\"protected.bin\" media-type=\"application/octet-stream\"/>" +
                "</manifest>" +
                "<spine><itemref idref=\"chapter\" properties=\"rendition:layout-reflowable\"/></spine>" +
                "</package>");
            WriteTextEntry(
                archive,
                "EPUB/chapter.xhtml",
                "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><h1>Chapter</h1><p>Body.</p></body></html>");
            WriteBytesEntry(archive, "EPUB/fonts/book.otf", new byte[] { 1, 2, 3, 4 });
            WriteBytesEntry(archive, "EPUB/protected.bin", new byte[] { 5, 6, 7, 8 });
            if (includeUnsafeEntry) {
                archive.CreateEntry("EPUB/");
                WriteTextEntry(archive, "../outside.xhtml", "<html><body>Outside</body></html>");
                WriteTextEntry(archive, "EPUB/../also-outside.xhtml", "<html><body>Also outside</body></html>");
            }
        }
        return output.ToArray();
    }

    private static byte[] BuildCaseSensitivePackage() {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteTextEntry(
                archive,
                "META-INF/container.xml",
                "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\">" +
                "<rootfiles><rootfile full-path=\"EPUB/package.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles>" +
                "</container>");
            WriteTextEntry(
                archive,
                "EPUB/package.opf",
                "<package version=\"3.0\" xmlns=\"http://www.idpf.org/2007/opf\">" +
                "<manifest>" +
                "<item id=\"Chapter\" href=\"Chapter.xhtml\" media-type=\"application/xhtml+xml\"/>" +
                "<item id=\"chapter\" href=\"chapter.xhtml\" media-type=\"application/xhtml+xml\"/>" +
                "</manifest><spine><itemref idref=\"Chapter\"/><itemref idref=\"chapter\"/></spine></package>");
            WriteTextEntry(archive, "EPUB/Chapter.xhtml", "<html><body><p>Upper body</p></body></html>");
            WriteTextEntry(archive, "EPUB/chapter.xhtml", "<html><body><p>Lower body</p></body></html>");
        }
        return output.ToArray();
    }

    private static byte[] BuildEpub2Package() {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteTextEntry(
                archive,
                "META-INF/container.xml",
                "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\">" +
                "<rootfiles><rootfile full-path=\"OPS/content.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles>" +
                "</container>");
            WriteTextEntry(
                archive,
                "OPS/content.opf",
                "<package version=\"2.0\" unique-identifier=\"book-id\" xmlns=\"http://www.idpf.org/2007/opf\">" +
                "<metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\">" +
                "<dc:identifier id=\"book-id\">urn:epub2:book</dc:identifier>" +
                "<dc:title>EPUB 2 contracts</dc:title><meta name=\"fixed-layout\" content=\"true\"/>" +
                "</metadata><manifest>" +
                "<item id=\"ncx\" href=\"toc.ncx\" media-type=\"application/x-dtbncx+xml\"/>" +
                "<item id=\"chapter\" href=\"chapter.xhtml\" media-type=\"application/xhtml+xml\"/>" +
                "</manifest><spine toc=\"ncx\"><itemref idref=\"chapter\"/></spine></package>");
            WriteTextEntry(
                archive,
                "OPS/toc.ncx",
                "<ncx xmlns=\"http://www.daisy.org/z3986/2005/ncx/\" version=\"2005-1\"><navMap>" +
                "<navPoint id=\"chapter\"><navLabel><text>NCX Chapter</text></navLabel>" +
                "<content src=\"chapter.xhtml\"/></navPoint></navMap></ncx>");
            WriteTextEntry(
                archive,
                "OPS/chapter.xhtml",
                "<html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Local title</title></head>" +
                "<body><p>EPUB 2 body</p></body></html>");
        }
        return output.ToArray();
    }

    private static void WriteTextEntry(
        ZipArchive archive,
        string path,
        string content,
        CompressionLevel compressionLevel = CompressionLevel.Optimal) {
        WriteBytesEntry(archive, path, Encoding.UTF8.GetBytes(content), compressionLevel);
    }

    private static void WriteBytesEntry(
        ZipArchive archive,
        string path,
        byte[] data,
        CompressionLevel compressionLevel = CompressionLevel.Optimal) {
        ZipArchiveEntry entry = archive.CreateEntry(path, compressionLevel);
        using Stream stream = entry.Open();
        stream.Write(data, 0, data.Length);
    }
}
