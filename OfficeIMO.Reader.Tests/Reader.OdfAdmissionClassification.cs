using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Testing;
using OfficeIMO.Reader.OpenDocument;
using OfficeIMO.Reader.Zip;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderOdfAdmissionClassificationTests {
    [Theory]
    [InlineData("integer", "direct")]
    [InlineData("integer", "lazy")]
    [InlineData("integer", "batch")]
    [InlineData("integer", "folder")]
    [InlineData("integer", "zip")]
    [InlineData("integer", "nested-zip")]
    [InlineData("ratio", "direct")]
    [InlineData("ratio", "lazy")]
    [InlineData("ratio", "batch")]
    [InlineData("ratio", "folder")]
    [InlineData("ratio", "zip")]
    [InlineData("ratio", "nested-zip")]
    [InlineData("encryption", "direct")]
    public async Task NativeAdmissionIsTerminalAcrossCanonicalReaderRoutes(string kind, string route) {
        byte[] bytes = Fixture(kind);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddOpenDocumentHandler().AddZipHandler().Build();
        var options = new ReaderOptions { OpenPassword = "fixture-password", Projection = route.Contains("zip") ? ReaderProjection.Full : ReaderProjection.TextAndMetadata };
        string directory = Path.Combine(Path.GetTempPath(), "odf-admission-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            async Task Read() {
                if (route == "direct") reader.ReadDocument(bytes, "source.odt", options);
                else if (route == "lazy") reader.Read(bytes, "source.odt", options).ToArray();
                else if (route.Contains("zip")) reader.ReadDocument(route == "zip" ? Wrap(bytes, "source.odt") : Wrap(Wrap(bytes, "source.odt"), "nested.zip"), "source.zip", options);
                else {
                    string path = Path.Combine(directory, "source.odt"); File.WriteAllBytes(path, bytes);
                    if (route == "batch") await reader.ReadDocumentsDetailedAsync(new[] { path }, options);
                    else reader.ReadFolder(directory, options: options).ToArray();
                }
            }
            if (kind == "integer") {
                ReaderResourceLimitException error = await Assert.ThrowsAsync<ReaderResourceLimitException>(Read);
                Assert.Equal("opendocument.MaximumDecodedCharacters", error.LimitName);
                Assert.Equal(16 * 1024 * 1024, error.Maximum); Assert.IsType<OfficeInspectionResourceLimitException>(error.InnerException);
            } else {
                ReaderAdmissionLimitException error = await Assert.ThrowsAsync<ReaderAdmissionLimitException>(Read);
                if (kind == "ratio") {
                    OfficePackageSecurityException native = Assert.IsType<OfficePackageSecurityException>(error.InnerException);
                    Assert.Equal("OpenDocument.MaxCompressionRatio", error.LimitName);
                    Assert.Equal(300d, error.Maximum); Assert.Equal(native.ObservedValue, error.Actual);
                } else {
                    OdfEncryptedPackageException native = Assert.IsType<OdfEncryptedPackageException>(error.InnerException);
                    Assert.Equal(OdfEncryptionFailureReason.ResourceLimitExceeded, native.Reason);
                    Assert.Equal("OpenDocument.Encryption.ResourceLimitExceeded", error.LimitName);
                    Assert.Null(error.Maximum); Assert.Null(error.Actual);
                }
            }
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("loaded")]
    [InlineData("lazy")]
    public void ExtractionBudgetRetainsExactConfiguredMaximum(string route) {
        OdtDocument document = OdtDocument.Create(); document.AddParagraph("Readable content");
        var format = new ReaderOpenDocumentOptions { MaxExtractedCharacters = 3 };
        var options = new ReaderOptions { Projection = ReaderProjection.TextAndMetadata };
        ReaderResourceLimitException error = Assert.Throws<ReaderResourceLimitException>(() => {
            if (route == "loaded") document.ReadDocument("source.odt", options, format);
            else new OfficeDocumentReaderBuilder().AddOpenDocumentHandler(format).Build().Read(document.ToBytes(), "source.odt", options).ToArray();
        });
        Assert.Equal("OpenDocument.MaxExtractedCharacters", error.LimitName); Assert.Equal(3, error.Maximum); Assert.Equal(3, format.MaxExtractedCharacters);
    }

    [Fact]
    public async Task MalformedOdfRemainsAnOrdinaryPerFileFailure() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".odt");
        try {
            File.WriteAllBytes(path, new byte[] { 1, 2, 3 });
            var reader = new OfficeDocumentReaderBuilder().AddOpenDocumentHandler().Build();
            ReaderDocumentReadOutcome outcome = Assert.Single(await reader.ReadDocumentsDetailedAsync(new[] { path }));
            Assert.IsType<InvalidDataException>(outcome.Error); Assert.Null(outcome.Document);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    private static byte[] Fixture(string kind) {
        OdtDocument document = OdtDocument.Create(); document.AddParagraph(kind == "ratio" ? new string('x', 1000000) : "Readable content");
        if (kind == "encryption") {
            byte[] encrypted = document.ToBytes(new OdfSaveOptions { Encryption = new OdfEncryptionOptions { Password = "fixture-password", IterationCount = 10000 } });
            return OdfTestPackageRewriter.Rewrite(encrypted, (name, bytes) => {
                if (name != "META-INF/manifest.xml") return bytes;
                XDocument manifest = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                XNamespace ns = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
                foreach (XElement derivation in manifest.Descendants(ns + "key-derivation")) derivation.SetAttributeValue(ns + "iteration-count", 10000000);
                return Encoding.UTF8.GetBytes(manifest.ToString());
            });
        }
        return kind == "ratio" ? document.ToBytes() : OdfTestPackageRewriter.Rewrite(document.ToBytes(), (name, bytes) => {
            if (name != "content.xml") return bytes;
            XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
            XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
            xml.Descendants(text + "p").Single().ReplaceNodes(new XElement(text + "s", new XAttribute(text + "c", 16777217)));
            return Encoding.UTF8.GetBytes(xml.ToString());
        });
    }

    private static byte[] Wrap(byte[] bytes, string name) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, true)) {
            using Stream entry = archive.CreateEntry(name, CompressionLevel.NoCompression).Open(); entry.Write(bytes, 0, bytes.Length);
        }
        return output.ToArray();
    }
}
