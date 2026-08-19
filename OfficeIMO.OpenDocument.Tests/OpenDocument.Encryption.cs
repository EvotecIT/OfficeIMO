using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentEncryptionTests {
    private const string Password = "OfficeIMO-ODF-test-2026";

    [Fact]
    public void Aes256PasswordEncryptionRoundTripsTextAndUsesStoredEncryptedEntries() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Encrypted café 中");

        byte[] encrypted = source.ToBytes(new OdfSaveOptions {
            Encryption = new OdfEncryptionOptions { Password = Password }
        });
        using (var stream = new MemoryStream(encrypted, writable: false))
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Read)) {
            ZipArchiveEntry encryptedContent = archive.GetEntry("content.xml")!;
            Assert.Equal(encryptedContent.Length, encryptedContent.CompressedLength);
            XDocument manifest = ReadXml(archive, "META-INF/manifest.xml");
            XNamespace ns = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
            XElement content = manifest.Root!.Elements(ns + "file-entry")
                .Single(element => (string?)element.Attribute(ns + "full-path") == "content.xml");
            Assert.NotNull(content.Element(ns + "encryption-data"));
            Assert.Equal("http://www.w3.org/2001/04/xmlenc#aes256-cbc",
                (string?)content.Descendants(ns + "algorithm").Single().Attribute(ns + "algorithm-name"));
        }

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(encrypted), new OdfLoadOptions { Password = Password });
        Assert.True(reopened.Security.SourceIsEncrypted);
        Assert.Equal("Encrypted café 中", reopened.ContentBlocks.Single().Paragraph!.Text);
    }

    [Fact]
    public void PasswordIsRequiredAndWrongPasswordIsClassified() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("secret");
        byte[] encrypted = source.ToBytes(new OdfSaveOptions {
            Encryption = new OdfEncryptionOptions { Password = Password }
        });
        string originalPackage = Convert.ToBase64String(encrypted);

        OdfEncryptedPackageException missing = Assert.Throws<OdfEncryptedPackageException>(() =>
            OdtDocument.Load(new MemoryStream(encrypted)));
        Assert.Equal(OdfEncryptionFailureReason.PasswordRequired, missing.Reason);

        OdfEncryptedPackageException wrong = Assert.Throws<OdfEncryptedPackageException>(() =>
            OdtDocument.Load(new MemoryStream(encrypted), new OdfLoadOptions { Password = "wrong-password" }));
        Assert.Equal(OdfEncryptionFailureReason.IncorrectPassword, wrong.Reason);
        Assert.False(string.IsNullOrEmpty(wrong.EntryPath));
        Assert.Equal(originalPackage, Convert.ToBase64String(encrypted));
    }

    [Fact]
    public void EncryptedSourceCannotBeSavedAsPlaintextWithoutExplicitRemoval() {
        OdtDocument original = OdtDocument.Create();
        original.AddParagraph("protected");
        byte[] encrypted = original.ToBytes(new OdfSaveOptions {
            Encryption = new OdfEncryptionOptions { Password = Password }
        });
        OdtDocument loaded = OdtDocument.Load(new MemoryStream(encrypted), new OdfLoadOptions { Password = Password });
        loaded.AddParagraph("changed");

        OdfEncryptedPackageException preservation = Assert.Throws<OdfEncryptedPackageException>(() => loaded.ToBytes());
        Assert.Equal(OdfEncryptionFailureReason.PreservationRequired, preservation.Reason);

        byte[] plaintext = loaded.ToBytes(new OdfSaveOptions { EncryptionHandling = OdfEncryptionHandling.Remove });
        OdtDocument reopened = OdtDocument.Load(new MemoryStream(plaintext));
        Assert.Equal(2, reopened.ContentBlocks.Count);
    }

    [Fact]
    public void EncryptionUsesFreshSaltAndIvForEverySave() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("same content");
        var options = new OdfSaveOptions { Encryption = new OdfEncryptionOptions { Password = Password } };

        byte[] first = source.ToBytes(options);
        byte[] second = source.ToBytes(options);

        Assert.NotEqual(Convert.ToBase64String(first), Convert.ToBase64String(second));
        Assert.Equal("same content", OdtDocument.Load(new MemoryStream(first), new OdfLoadOptions { Password = Password })
            .ContentBlocks.Single().Paragraph!.Text);
        Assert.Equal("same content", OdtDocument.Load(new MemoryStream(second), new OdfLoadOptions { Password = Password })
            .ContentBlocks.Single().Paragraph!.Text);
    }

    [Fact]
    public void EncryptionRoundTripsSpreadsheetAndPresentationPackages() {
        OdsDocument spreadsheet = OdsDocument.Create();
        spreadsheet.AddSheet("Secure").Cell(0, 0).SetString("sheet secret");
        byte[] ods = spreadsheet.ToBytes(new OdfSaveOptions { Encryption = new OdfEncryptionOptions { Password = Password } });
        Assert.Equal("sheet secret", OdsDocument.Load(new MemoryStream(ods), new OdfLoadOptions { Password = Password })
            .Sheets[0].Cell(0, 0).Value.DisplayText);

        OdpPresentation presentation = OdpPresentation.Create();
        presentation.AddSlide().AddTextBox(OdfRect.FromCentimeters(1, 1, 5, 2), "slide secret");
        byte[] odp = presentation.ToBytes(new OdfSaveOptions { Encryption = new OdfEncryptionOptions { Password = Password } });
        OdpPresentation reopened = OdpPresentation.Load(new MemoryStream(odp), new OdfLoadOptions { Password = Password });
        Assert.Single(reopened.Slides);
    }

    [Fact]
    public void LibreOfficeEncryptedProducerFixtureIsHashPinnedAndReadable() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Encryption",
            "libreoffice-24.2-aes256.odt");
        byte[] bytes = File.ReadAllBytes(path);
        using (SHA256 sha = SHA256.Create()) {
            Assert.Equal("c6b6f5feb4c3122528eb6d12f615456deab7f23f0a48e8d9482f8706fccd870b",
                BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", string.Empty).ToLowerInvariant());
        }

        OdtDocument document = OdtDocument.Load(new MemoryStream(bytes),
            new OdfLoadOptions { Password = Password });
        Assert.Contains(document.ContentBlocks, block =>
            block.Paragraph?.Text.IndexOf("LibreOffice encrypted producer fixture", StringComparison.Ordinal) >= 0);

        OdfEncryptedPackageException wrong = Assert.Throws<OdfEncryptedPackageException>(() =>
            OdtDocument.Load(new MemoryStream(bytes), new OdfLoadOptions { Password = "wrong-password" }));
        Assert.Equal(OdfEncryptionFailureReason.IncorrectPassword, wrong.Reason);
    }

    [Fact]
    public void DecryptedEntrySizeIsCheckedBeforeLargePlaintextIsExposed() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph(new string('A', 200000));
        byte[] encrypted = source.ToBytes(new OdfSaveOptions {
            Encryption = new OdfEncryptionOptions { Password = Password }
        });

        OdfEncryptedPackageException limit = Assert.Throws<OdfEncryptedPackageException>(() =>
            OdtDocument.Load(new MemoryStream(encrypted), new OdfLoadOptions {
                Password = Password,
                MaxEntryUncompressedBytes = 64L * 1024L
            }));

        Assert.Equal(OdfEncryptionFailureReason.ResourceLimitExceeded, limit.Reason);
        Assert.Equal("content.xml", limit.EntryPath);
    }

    [Fact]
    public void AggregateKdfWorkIsPreflightedBeforeAnyEntryIsDecrypted() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("bounded KDF work");
        byte[] encrypted = source.ToBytes(new OdfSaveOptions {
            Encryption = new OdfEncryptionOptions { Password = Password, IterationCount = 100_000 }
        });

        OdfEncryptedPackageException limit = Assert.Throws<OdfEncryptedPackageException>(() =>
            OdtDocument.Load(new MemoryStream(encrypted), new OdfLoadOptions {
                Password = Password,
                MaxTotalKdfIterations = 100_000
            }));

        Assert.Equal(OdfEncryptionFailureReason.ResourceLimitExceeded, limit.Reason);
        Assert.False(string.IsNullOrEmpty(limit.EntryPath));
    }

    private static XDocument ReadXml(ZipArchive archive, string path) {
        using Stream stream = archive.GetEntry(path)!.Open();
        return XDocument.Load(stream);
    }
}
