using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeIMO.Epub;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubWriterPackagePreservationContracts {
    private static readonly XNamespace Opf = "http://www.idpf.org/2007/opf";
    private static readonly XNamespace Container = "urn:oasis:names:tc:opendocument:xmlns:container";

    [Fact]
    public async Task OriginalArchive_EnforcesPhysicalDirectoryCountBeforeSaving() {
        byte[] source = Rewrite(EpubWritingContracts.CreateBook().Write().Bytes,
            new Dictionary<string, byte[]> { ["EPUB/"] = Array.Empty<byte>(), ["META-INF/"] = Array.Empty<byte>() });
        using var archive = new ZipArchive(new MemoryStream(source));
        EpubPublication book = EpubPublication.Load(new MemoryStream(source));
        Assert.Equal(source, book.Write(new EpubWriteOptions { MaxEntries = archive.Entries.Count }).Bytes);
        var limits = new EpubWriteOptions { MaxEntries = archive.Entries.Count - 1 };
        using var destination = new MemoryStream(new byte[] { 1, 2, 3 }, true);
        Assert.Throws<InvalidDataException>(() => book.Save(destination, limits));
        await Assert.ThrowsAsync<InvalidDataException>(() => book.SaveAsync(destination, limits));
        Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
        Assert.True(book.Write(new EpubWriteOptions { MaxEntries = archive.Entries.Count - 2, CompressEntries = false }).Bytes.Length > 0);
    }

    [Theory]
    [InlineData("manifest")]
    [InlineData("metadata")]
    [InlineData("content")]
    [InlineData("rootfile")]
    [InlineData("unreadable")]
    public void Removal_ProtectsAlternateRootfilesBeforeMutation(string reference) {
        EpubPublication book = AlternatePublication(reference);
        string id = reference == "rootfile" ? "alternate" : "shared";
        byte[] before = book.Write().Bytes;
        Assert.ThrowsAny<InvalidOperationException>(() => book.RemoveResource(id));
        Assert.Contains(book.Manifest, item => item.Id == id);
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void Removal_AllowsUnreferencedPayloadAndPreservesAlternatePackageBytes() {
        EpubPublication book = AlternatePublication("none");
        byte[] alternate = book.GetResourceBytes("alternate");
        book.RemoveResource("shared");
        EpubPublication output = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.DoesNotContain("EPUB/shared.dat", output.EntryPaths);
        Assert.Equal(alternate, output.GetResourceBytes("alternate"));
    }

    [Theory]
    [InlineData("xhtml")]
    [InlineData("svg")]
    [InlineData("ncx")]
    public void RaisedEntryBound_AppliesToRetainedXmlInspectionAndWrite(string kind) {
        EpubPublication book = EpubWritingContracts.CreateBook(kind == "ncx" ? EpubVersion.Epub2 : EpubVersion.Epub3);
        string id = kind == "ncx" ? "navigation" : "first";
        if (kind == "svg") {
            id = "illustration";
            book.AddResource(id, "EPUB/illustration.svg", "image/svg+xml", Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg'><title>Image</title></svg>"));
        }
        XDocument content = book.GetContentXml(id);
        content.Root!.Add(new XComment(new string('a', 73 * 1024 * 1024)));
        byte[] payload = Encoding.UTF8.GetBytes(content.ToString(SaveOptions.DisableFormatting));
        string path = book.Manifest.Single(item => item.Id == id).Reference.ContainerPath!;
        byte[] source = Rewrite(book.Write().Bytes, new Dictionary<string, byte[]> { [path] = payload });
        book = EpubPublication.Load(new MemoryStream(source), new EpubPublicationLoadOptions { MaxEntryBytes = payload.LongLength + 4096 });
        Assert.Equal(source, book.Write().Bytes);
        Assert.Equal(content.Root.Name, book.GetContentXml(id).Root!.Name);
        if (kind == "ncx") {
            book.Identifier = "urn:example:changed";
            book.SetNavigation(new[] { new EpubNavigationEntry("First", "EPUB/first.xhtml") });
            book.AddChapter("third", "EPUB/third.xhtml", "Third", "<p>Third</p>");
            Assert.True(book.Write().Bytes.Length > 0);
        } else {
            content.Root.SetAttributeValue("data-edited", "yes");
            book.SetContentXml(id, content);
            Assert.True(book.Write().Bytes.Length > 0);
        }
    }

    [Theory]
    [InlineData("mimetype")]
    [InlineData("META-INF/container.xml")]
    [InlineData("META-INF/encryption.xml")]
    [InlineData("EPUB/package.opf")]
    [InlineData("EPUB/alternate.opf")]
    public void ControlResourceMutations_RejectRawReplacementAndRemovalBeforeMutation(string path) {
        EpubPublication book = path == "EPUB/alternate.opf" ? AlternatePublication("none") : ControlPublication(path);
        string id = path == "EPUB/alternate.opf" ? "alternate" : "control";
        byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidOperationException>(() => book.UpdateResource(id, Array.Empty<byte>()));
        Assert.Throws<InvalidOperationException>(() => book.RemoveResource(id));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public async Task ManifestedSignatureRemoval_RequiresExplicitPolicyAtSyncAndAsyncSave() {
        EpubPublication book = ControlPublication("META-INF/signatures.xml");
        Assert.Throws<InvalidOperationException>(() => book.UpdateResource("control", Array.Empty<byte>()));
        book.RemoveResource("control");
        using var destination = new MemoryStream(new byte[] { 1, 2, 3 }, true);
        Assert.Throws<InvalidOperationException>(() => book.Save(destination));
        await Assert.ThrowsAsync<InvalidOperationException>(() => book.SaveAsync(destination));
        Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
        EpubWriteResult result = book.Write(new EpubWriteOptions { RemoveInvalidatedSignatures = true });
        using var archive = new ZipArchive(new MemoryStream(result.Bytes));
        Assert.Null(archive.GetEntry("META-INF/signatures.xml"));
        Assert.Contains(result.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "EPUB_WRITE_SIGNATURE_REMOVED");
        Assert.True(result.HasLoss);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RemovalBudget_CreditsOnlyPayloadsThatWillActuallyBeReleased(bool retainedAlias) {
        EpubPublication book = EpubPublication.Create(new string('>', 600));
        book.AddChapter("first", "EPUB/first.xhtml", "First", "<p>First</p>");
        book.AddResource("extra", "EPUB/extra.dat", "application/octet-stream", new byte[4096]);
        byte[] source = book.Write().Bytes;
        XDocument package = XDocument.Parse(Encoding.UTF8.GetString(ReadEntry(source, book.PackagePath)));
        if (retainedAlias) package.Root!.Element(Opf + "manifest")!.Add(new XElement(Opf + "item", new XAttribute("id", "alias"),
            new XAttribute("href", "extra.dat"), new XAttribute("media-type", "application/octet-stream")));
        source = Rewrite(source, new Dictionary<string, byte[]> { [book.PackagePath] = Encoding.UTF8.GetBytes(package.ToString().Replace("&gt;", ">")) });
        using var archive = new ZipArchive(new MemoryStream(source));
        long maximum = archive.Entries.Sum(entry => entry.Length);
        book = EpubPublication.Load(new MemoryStream(source), new EpubPublicationLoadOptions { MaxExpandedBytes = maximum });
        if (retainedAlias) {
            Assert.Throws<InvalidDataException>(() => book.RemoveResource("extra"));
            Assert.Equal(source, book.Write().Bytes);
        } else {
            book.RemoveResource("extra");
            using var output = new ZipArchive(new MemoryStream(book.Write().Bytes));
            Assert.Null(output.GetEntry("EPUB/extra.dat"));
            Assert.True(output.Entries.Sum(entry => entry.Length) <= maximum);
            Assert.Equal(new string('>', 600), book.Title);
        }
    }

    private static EpubPublication ControlPublication(string path) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        if (path == "META-INF/encryption.xml") book.AddResource("protected", "EPUB/protected.dat", "application/octet-stream", new byte[] { 42 });
        byte[] source = book.Write().Bytes;
        XDocument package = XDocument.Parse(Encoding.UTF8.GetString(ReadEntry(source, book.PackagePath)));
        package.Root!.Element(Opf + "manifest")!.Add(new XElement(Opf + "item", new XAttribute("id", "control"),
            new XAttribute("href", path.StartsWith("EPUB/", StringComparison.Ordinal) ? path.Substring(5) : "../" + path), new XAttribute("media-type", "application/xml")));
        var replacements = new Dictionary<string, byte[]> { [book.PackagePath] = Encoding.UTF8.GetBytes(package.ToString()) };
        if (path == "META-INF/signatures.xml") replacements[path] = Encoding.UTF8.GetBytes("<signatures xmlns='urn:oasis:names:tc:opendocument:xmlns:container'/>");
        if (path == "META-INF/encryption.xml") replacements[path] = Encoding.UTF8.GetBytes(
            "<encryption xmlns='urn:oasis:names:tc:opendocument:xmlns:container' xmlns:e='http://www.w3.org/2001/04/xmlenc#'><e:EncryptedData>" +
            "<e:EncryptionMethod Algorithm='http://www.idpf.org/2008/embedding'/><e:CipherData><e:CipherReference URI='EPUB/protected.dat'/></e:CipherData></e:EncryptedData></encryption>");
        return EpubPublication.Load(new MemoryStream(Rewrite(source, replacements)));
    }

    private static EpubPublication AlternatePublication(string reference) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        book.AddResource("shared", "EPUB/shared.dat", "application/octet-stream", new byte[] { 42 });
        XDocument alternate = book.GetPackageXml();
        alternate.Descendants(Opf + "item").Single(item => (string?)item.Attribute("id") == "shared").Remove();
        alternate.Root!.Element(Opf + "manifest")!.Add(new XElement(Opf + "item", new XAttribute("id", "alternate-content"),
            new XAttribute("href", "alternate.xhtml"), new XAttribute("media-type", "application/xhtml+xml")));
        if (reference == "manifest") alternate.Root.Element(Opf + "manifest")!.Add(new XElement(Opf + "item",
            new XAttribute("id", "shared"), new XAttribute("href", "shared.dat"), new XAttribute("media-type", "application/octet-stream")));
        if (reference == "metadata") alternate.Root.Element(Opf + "metadata")!.Add(new XElement(Opf + "link",
            new XAttribute("rel", "record"), new XAttribute("href", "shared.dat")));
        book.AddResource("alternate", "EPUB/alternate.opf", "application/oebps-package+xml", reference == "unreadable"
            ? Encoding.UTF8.GetBytes("unreadable") : Encoding.UTF8.GetBytes(alternate.ToString()));
        byte[] source = book.Write().Bytes;
        XDocument container = XDocument.Parse(Encoding.UTF8.GetString(ReadEntry(source, "META-INF/container.xml")));
        container.Root!.Element(Container + "rootfiles")!.Add(new XElement(Container + "rootfile",
            new XAttribute("full-path", "EPUB/alternate.opf"), new XAttribute("media-type", "application/oebps-package+xml")));
        source = Rewrite(source, new Dictionary<string, byte[]> {
            ["META-INF/container.xml"] = Encoding.UTF8.GetBytes(container.ToString()),
            ["EPUB/alternate.xhtml"] = Encoding.UTF8.GetBytes("<html xmlns='http://www.w3.org/1999/xhtml'><head><title>Alternate</title>" +
                (reference == "content" ? "<base href='sub/'/>" : "") + "</head><body>" +
                (reference == "content" ? "<a href='../shared.dat'>Shared</a>" : "<p>Alternate</p>") + "</body></html>")
        });
        return EpubPublication.Load(new MemoryStream(source));
    }

    private static byte[] ReadEntry(byte[] source, string path) {
        using var archive = new ZipArchive(new MemoryStream(source));
        using Stream input = archive.GetEntry(path)!.Open();
        using var output = new MemoryStream(); input.CopyTo(output); return output.ToArray();
    }

    private static byte[] Rewrite(byte[] source, Dictionary<string, byte[]> replacements) {
        using var input = new ZipArchive(new MemoryStream(source));
        var entries = input.Entries.ToDictionary(entry => entry.FullName, entry => ReadEntry(source, entry.FullName), StringComparer.Ordinal);
        foreach (var pair in replacements) entries[pair.Key] = pair.Value;
        return OfficeProvenanceZipWriter.Write(entries.OrderBy(pair => pair.Key == "mimetype" ? 0 : 1)
            .Select(pair => new OfficeProvenanceZipWriteEntry(pair.Key, pair.Value.LongLength, pair.Key != "mimetype",
                new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero), 0, 0, Array.Empty<byte>(), Array.Empty<byte>(), Array.Empty<byte>(),
                () => new MemoryStream(pair.Value, false))).ToArray(), 256L * 1024 * 1024);
    }
}
