using OfficeIMO.Epub;
using OfficeIMO.Html;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookDeliveryTests {
    [Fact]
    public void DeliveryBindsExactExportedMetadataAndPublicationWithoutEditorialHistory() {
        var project = BookProject.Create("First edition", "pl");
        project.CreateRevision("Private editorial baseline");
        project.SetMetadata("Published edition", "pl", "Author");
        project.Publication.AddIdentifier("isbn", new EpubIdentifierMetadata { Value = "978-0-306-40615-7", Kind = EpubIdentifierKind.Isbn13 });
        project.Publication.AddContributor("translator", new EpubContributorMetadata { Name = "Translator", MarcRoles = ["trl"], Language = "pl" });
        byte[] before = project.ToProjectBytes();
        var options = new EpubWriteOptions { ModifiedAt = new DateTimeOffset(2026, 10, 5, 12, 0, 0, TimeSpan.Zero) };
        byte[] delivery = project.ToDeliveryBytes(options);
        Assert.Equal(delivery, project.ToDeliveryBytes(options));
        var entries = Entries(delivery);
        Assert.Equal(new[] { "manifest.json", "package.opf", "publication.epub" }, entries.Keys.Order());
        Assert.Equal(project.Export(options).Bytes, entries["publication.epub"]);
        var epubEntries = Entries(entries["publication.epub"]);
        Assert.Equal(epubEntries[project.Publication.PackagePath], entries["package.opf"]);
        var package = XDocument.Load(new MemoryStream(entries["package.opf"]));
        XNamespace opf = "http://www.idpf.org/2007/opf";
        Assert.Equal("2026-10-05T12:00:00Z", package.Descendants(opf + "meta").Single(e => (string?)e.Attribute("property") == "dcterms:modified").Value);
        Assert.Contains("9780306406157", package.ToString());
        Assert.Contains("trl", package.ToString());
        using var json = JsonDocument.Parse(entries["manifest.json"]);
        var manifest = json.RootElement;
        Assert.Equal("OfficeIMO.BookDelivery", manifest.GetProperty("Format").GetString());
        Assert.Equal(1, manifest.GetProperty("Version").GetInt32());
        Assert.Equal(project.Publication.PackagePath, manifest.GetProperty("PackagePath").GetString());
        Assert.Equal("passed", manifest.GetProperty("NativeWriterValidation").GetString());
        Assert.Equal("not-performed", manifest.GetProperty("IndependentValidation").GetString());
        Assert.Equal("not-performed", manifest.GetProperty("RetailerAcceptance").GetString());
        Assert.Equal(2, manifest.GetProperty("Files").GetArrayLength());
        foreach (var file in manifest.GetProperty("Files").EnumerateArray()) {
            byte[] payload = entries[file.GetProperty("Name").GetString()!];
            Assert.Equal(payload.LongLength, file.GetProperty("Bytes").GetInt64());
            Assert.Equal(Convert.ToHexString(SHA256.HashData(payload)), file.GetProperty("Sha256").GetString());
        }
        Assert.Equal(before, project.ToProjectBytes());
        project.Undo(); Assert.Equal("First edition", project.Publication.Title);
    }

    [Fact]
    public void ImportReviewStillGatesDeliveryAndAcceptedFindingsRemainStructured() {
        var project = BookProject.FromImport(EpubManuscript.ImportHtml(HtmlConversionDocument.Parse(
            "<title>Book</title><h1>One</h1><p>Text</p><script>ignored()</script>")));
        Assert.Throws<InvalidOperationException>(() => project.ToDeliveryBytes());
        project.AcknowledgeImportLoss();
        using var json = JsonDocument.Parse(Entries(project.ToDeliveryBytes())["manifest.json"]);
        Assert.True(json.RootElement.GetProperty("ImportLossAcknowledged").GetBoolean());
        var findings = json.RootElement.GetProperty("ImportDiagnostics").EnumerateArray().ToArray();
        Assert.Contains(findings, finding => finding.GetProperty("Code").GetString() == "EPUB_IMPORT_ACTIVE_CONTENT_OMITTED" &&
            finding.GetProperty("LossKind").GetString() == "Omission");
    }

    [Fact]
    public void InvalidatedSignaturesRequireExplicitRemovalAndRemainReported() {
        using var source = new MemoryStream();
        byte[] original = BookProject.Create("Signed title").Export().Bytes;
        source.Write(original);
        using (var archive = new ZipArchive(source, ZipArchiveMode.Update, true)) {
            using var writer = new StreamWriter(archive.CreateEntry("META-INF/signatures.xml").Open());
            writer.Write("<signatures xmlns='urn:oasis:names:tc:opendocument:xmlns:container'/>");
        }
        var project = BookProject.FromEpub(source.ToArray());
        project.Publication.Title = "Edited title";
        Assert.Throws<InvalidOperationException>(() => project.ToDeliveryBytes());
        var delivery = Entries(project.ToDeliveryBytes(new EpubWriteOptions { RemoveInvalidatedSignatures = true }));
        Assert.DoesNotContain("META-INF/signatures.xml", Entries(delivery["publication.epub"]).Keys);
        using var json = JsonDocument.Parse(delivery["manifest.json"]);
        Assert.Contains(json.RootElement.GetProperty("WriterDiagnostics").EnumerateArray(),
            finding => finding.GetProperty("LossKind").GetString() == "Omission");
        // Export policy does not silently remove the signature from the editable source model.
        Assert.Throws<InvalidOperationException>(() => project.Export());
    }

    [Fact]
    public void OutputLimitsAndCancellationDoNotMutateProjectOrCallerOptions() {
        var project = BookProject.Create("Book");
        byte[] before = project.ToProjectBytes();
        var options = new EpubWriteOptions { MaxOutputBytes = 256L * 1024 * 1024 };
        byte[] complete = project.ToDeliveryBytes(options);
        Assert.Equal(256L * 1024 * 1024, options.MaxOutputBytes);
        // This bound permits the EPUB itself but fails while packaging the delivery.
        Assert.Throws<InvalidDataException>(() => project.ToDeliveryBytes(options, complete.Length - 1));
        Assert.Throws<InvalidDataException>(() => project.ToDeliveryBytes(new EpubWriteOptions { MaxOutputBytes = 16 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => project.ToDeliveryBytes(maximumOutputBytes: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => project.ToDeliveryBytes(maximumOutputBytes: BookProject.MaximumDeliveryBytes + 1));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => project.ToDeliveryBytes(cancellationToken: cancellation.Token));
        Assert.Equal(before, project.ToProjectBytes());
    }

    private static Dictionary<string, byte[]> Entries(byte[] bytes) {
        using var input = new MemoryStream(bytes, false);
        using var archive = new ZipArchive(input, ZipArchiveMode.Read);
        return archive.Entries.ToDictionary(entry => entry.FullName, entry => {
            using var stream = entry.Open(); using var output = new MemoryStream();
            stream.CopyTo(output); return output.ToArray();
        });
    }
}
