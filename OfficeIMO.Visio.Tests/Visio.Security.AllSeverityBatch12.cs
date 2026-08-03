using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioAllSeverityBatch12SecurityTests {
    [Fact]
    public async Task LoadApisEnforceInputBudgetForNonSeekableStreams() {
        VisioDocument source = VisioDocument.Create();
        source.AddPage("Page-1");
        byte[] package = source.ToBytes();
        var restrictive = new VisioLoadOptions {
            MaxInputBytes = package.Length - 1L
        };

        using var stream = new Batch12NonSeekableStream(package);
        Assert.Throws<InvalidDataException>(() =>
            VisioDocument.Load(stream, restrictive));

        using var asyncStream = new Batch12NonSeekableStream(package);
        await Assert.ThrowsAsync<InvalidDataException>(() =>
            VisioDocument.LoadAsync(
                asyncStream, restrictive, CancellationToken.None));

        using var validStream = new MemoryStream(package, writable: false);
        VisioDocument valid = VisioDocument.Load(validStream,
            new VisioLoadOptions { MaxInputBytes = package.Length });
        Assert.Single(valid.Pages);
    }

    [Fact]
    public async Task PathLoadApisEnforceTheSameInputBudgetAsStreamLoads() {
        string path = Path.Combine(
            Path.GetTempPath(),
            "OfficeIMO.Visio.LoadBudget." + Guid.NewGuid().ToString("N") + ".vsdx");
        try {
            VisioDocument source = VisioDocument.Create(path);
            source.AddPage("Page-1");
            source.Save();
            var restrictive = new VisioLoadOptions {
                MaxInputBytes = new FileInfo(path).Length - 1L
            };

            Assert.Throws<InvalidDataException>(() =>
                VisioDocument.Load(path, restrictive));
            await Assert.ThrowsAsync<InvalidDataException>(() =>
                VisioDocument.LoadAsync(path, restrictive, default));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ExtractMastersContainsHostileIdsWithinOutputFolder() {
        string root = Path.Combine(Path.GetTempPath(),
            "officeimo-visio-assets-" + Guid.NewGuid().ToString("N"));
        string packagePath = Path.Combine(root, "source.vsdx");
        string output = Path.Combine(root, "output");
        string escaped = Path.Combine(root, "escaped-Danger.xml");
        Directory.CreateDirectory(root);
        try {
            CreateMasterPackage(packagePath, "../escaped", "Danger");

            VisioAssets.ExtractMasters(packagePath, output);

            string extracted = Assert.Single(Directory.GetFiles(output));
            Assert.StartsWith(Path.GetFullPath(output)
                    + Path.DirectorySeparatorChar,
                Path.GetFullPath(extracted),
                StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(escaped));
            Assert.Contains("escaped-", Path.GetFileName(extracted),
                StringComparison.Ordinal);
#if NET6_0_OR_GREATER
            if (!OperatingSystem.IsWindows()) {
                File.Delete(extracted);
                File.WriteAllText(escaped, "outside-safe");
                File.CreateSymbolicLink(extracted, escaped);

                Assert.Throws<IOException>(() =>
                    VisioAssets.ExtractMasters(packagePath, output));
                Assert.Equal("outside-safe", File.ReadAllText(escaped));
            }
#endif
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
    }

    private static void CreateMasterPackage(
        string path,
        string id,
        string name) {
        using ZipArchive archive = ZipFile.Open(path,
            ZipArchiveMode.Create);
        WriteEntry(archive,
            "visio/masters/masters.xml",
            $"<Masters xmlns=\"http://schemas.microsoft.com/office/visio/2012/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><Master ID=\"{id}\" NameU=\"{name}\"><Rel r:id=\"rId1\" /></Master></Masters>");
        WriteEntry(archive,
            "visio/masters/_rels/masters.xml.rels",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"urn:officeimo:test\" Target=\"master1.xml\" /></Relationships>");
        WriteEntry(archive,
            "visio/masters/master1.xml",
            "<MasterContents xmlns=\"http://schemas.microsoft.com/office/visio/2012/main\" />");
    }

    private static void WriteEntry(
        ZipArchive archive,
        string name,
        string content) {
        ZipArchiveEntry entry = archive.CreateEntry(name);
        using Stream stream = entry.Open();
        using var writer = new StreamWriter(stream,
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(content);
    }

    private sealed class Batch12NonSeekableStream : Stream {
        private readonly MemoryStream _inner;

        internal Batch12NonSeekableStream(byte[] bytes) {
            _inner = new MemoryStream(bytes, writable: false);
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) =>
            _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) =>
            throw new NotSupportedException();
        public override void SetLength(long value) =>
            throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) =>
            throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
