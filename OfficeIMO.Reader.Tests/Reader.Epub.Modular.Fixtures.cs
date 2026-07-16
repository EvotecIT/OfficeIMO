using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Tests;

public sealed partial class ReaderEpubModularTests {
    private static void BuildEpubWithSpine(
        string epubPath,
        string secondTitleXml = "Second",
        bool includePackageDiagnostics = false) {
        using var fs = new FileStream(epubPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(fs, ZipArchiveMode.Create, leaveOpen: false);

        WriteTextEntry(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
        WriteTextEntry(archive, "META-INF/container.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\">" +
            "<rootfiles><rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles>" +
            "</container>");
        if (includePackageDiagnostics) {
            WriteTextEntry(
                archive,
                "META-INF/encryption.xml",
                "<encryption xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\" xmlns:enc=\"http://www.w3.org/2001/04/xmlenc#\">" +
                "<enc:EncryptedData><enc:EncryptionMethod Algorithm=\"urn:unsupported\"/><enc:CipherData>" +
                "<enc:CipherReference URI=\"OEBPS/protected.bin\"/></enc:CipherData></enc:EncryptedData></encryption>");
        }

        WriteTextEntry(archive, "OEBPS/content.opf",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<package version=\"3.0\" unique-identifier=\"bookid\" xmlns=\"http://www.idpf.org/2007/opf\">" +
            "<metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\">" +
            "<dc:title>Demo Book</dc:title><dc:creator>OfficeIMO Team</dc:creator><dc:language>en</dc:language><dc:identifier id=\"bookid\">author-1</dc:identifier>" +
            (includePackageDiagnostics ? "<meta property=\"rendition:layout\">pre-paginated</meta>" : string.Empty) +
            "</metadata>" +
            "<manifest>" +
            "<item id=\"nav\" href=\"nav.xhtml\" media-type=\"application/xhtml+xml\" properties=\"nav\"/>" +
            "<item id=\"ncx\" href=\"toc.ncx\" media-type=\"application/x-dtbncx+xml\"/>" +
            "<item id=\"ch1\" href=\"chapter1.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "<item id=\"ch2\" href=\"chapter2.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "<item id=\"cover\" href=\"images/cover.png\" media-type=\"image/png\" properties=\"cover-image\"/>" +
            (includePackageDiagnostics ? "<item id=\"protected\" href=\"protected.bin\" media-type=\"application/octet-stream\"/>" : string.Empty) +
            "</manifest>" +
            "<spine toc=\"ncx\"><itemref idref=\"ch2\"/><itemref idref=\"ch1\"/></spine>" +
            "</package>");

        WriteTextEntry(archive, "OEBPS/nav.xhtml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><nav epub:type=\"toc\" xmlns:epub=\"http://www.idpf.org/2007/ops\"><ol>" +
            "<li><a href=\"chapter2.xhtml\">" + secondTitleXml + "</a></li>" +
            "<li><a href=\"chapter1.xhtml\">First</a></li>" +
            "</ol></nav></body></html>");

        WriteTextEntry(archive, "OEBPS/toc.ncx",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<ncx xmlns=\"http://www.daisy.org/z3986/2005/ncx/\" version=\"2005-1\"><navMap>" +
            "<navPoint id=\"n1\"><navLabel><text>" + secondTitleXml + "</text></navLabel><content src=\"chapter2.xhtml\"/></navPoint>" +
            "<navPoint id=\"n2\"><navLabel><text>First</text></navLabel><content src=\"chapter1.xhtml\"/></navPoint>" +
            "</navMap></ncx>");

        WriteTextEntry(archive, "OEBPS/chapter1.xhtml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Local One</title></head><body><h1>One</h1><p>First chapter text.</p>" +
            "<p><a href=\"chapter2.xhtml#details\">next chapter</a></p>" +
            "<img src=\"data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==\" alt=\"Inline\"/></body></html>");

        WriteTextEntry(archive, "OEBPS/chapter2.xhtml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Local Two</title></head><body><h1>Two</h1><p>Second chapter text. <a href=\"https://example.test/chapter-two\">details</a></p>" +
            "<ul><li>EPUB list item</li></ul>" +
            "<table><tr><th>Name</th><th>Qty</th></tr><tr><td>Chapter</td><td>2</td></tr></table>" +
            "<img src=\"data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==\" alt=\"Inline\"/>" +
            "<img src=\"images/cover.png\" alt=\"Cover\"/></body></html>");

        WriteBinaryEntry(archive, "OEBPS/images/cover.png", new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 });
        if (includePackageDiagnostics) {
            WriteBinaryEntry(archive, "OEBPS/protected.bin", new byte[] { 1, 2, 3, 4 });
        }
    }

    private static void BuildImageOnlyEpub(string epubPath) {
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
            "<package version=\"3.0\" xmlns=\"http://www.idpf.org/2007/opf\">" +
            "<manifest><item id=\"cover-page\" href=\"cover.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "<item id=\"cover\" href=\"images/cover.png\" media-type=\"image/png\" properties=\"cover-image\"/></manifest>" +
            "<spine><itemref idref=\"cover-page\"/></spine></package>");
        WriteTextEntry(archive, "OEBPS/cover.xhtml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Cover</title></head>" +
            "<body><img src=\"images/cover.png\" alt=\"Cover\"/></body></html>");
        WriteBinaryEntry(archive, "OEBPS/images/cover.png", new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 });
    }

    private static void BuildEpubWithEncryptedChapter(string epubPath) {
        using var fs = new FileStream(epubPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(fs, ZipArchiveMode.Create, leaveOpen: false);

        WriteTextEntry(
            archive,
            "META-INF/container.xml",
            "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\">" +
            "<rootfiles><rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles>" +
            "</container>");
        WriteTextEntry(
            archive,
            "META-INF/encryption.xml",
            "<encryption xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\" xmlns:enc=\"http://www.w3.org/2001/04/xmlenc#\">" +
            "<enc:EncryptedData><enc:EncryptionMethod Algorithm=\"urn:unsupported\"/><enc:CipherData>" +
            "<enc:CipherReference URI=\"OEBPS/locked.xhtml\"/></enc:CipherData></enc:EncryptedData></encryption>");
        WriteTextEntry(
            archive,
            "OEBPS/content.opf",
            "<package version=\"3.0\" xmlns=\"http://www.idpf.org/2007/opf\"><manifest>" +
            "<item id=\"locked\" href=\"locked.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "<item id=\"open\" href=\"open.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "</manifest><spine><itemref idref=\"locked\"/><itemref idref=\"open\"/></spine></package>");
        WriteTextEntry(
            archive,
            "OEBPS/./locked.xhtml",
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><p>Locked body.</p></body></html>");
        WriteTextEntry(
            archive,
            "OEBPS/open.xhtml",
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><p>Open body.</p></body></html>");
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

    private static void WriteBinaryEntry(ZipArchive archive, string path, byte[] content) {
        ZipArchiveEntry entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using Stream stream = entry.Open();
        stream.Write(content, 0, content.Length);
    }

    private static string ComputeSha256Hex(byte[] content) {
        byte[] hash;
        using (SHA256 sha = SHA256.Create()) {
            hash = sha.ComputeHash(content);
        }

        var result = new StringBuilder(hash.Length * 2);
        foreach (byte value in hash) {
            result.Append(value.ToString("x2"));
        }
        return result.ToString();
    }
}
