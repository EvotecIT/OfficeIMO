using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Packages separately encoded retained HTML surfaces with a deterministic manifest.</summary>
public static class HtmlRenderResultArchiveExtensions {
    private static readonly DateTimeOffset StableTimestamp =
        new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

    /// <summary>
    /// Encodes the retained PNG or SVG surfaces and packages them with <c>manifest.json</c>
    /// under a bounded packaging deadline derived from the retained request options.
    /// </summary>
    public static HtmlRenderArchiveResult ExportArchive(this HtmlRenderResult result,
        HtmlRenderArchiveOptions? archiveOptions = null,
        CancellationToken cancellationToken = default) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        cancellationToken.ThrowIfCancellationRequested();
        if (result.Request.Encoder != HtmlRenderEncoder.Png && result.Request.Encoder != HtmlRenderEncoder.Svg) {
            throw new InvalidOperationException("HTML render archives require an explicit PNG or SVG render encoder.");
        }

        HtmlRenderArchiveOptions packaging = archiveOptions?.Clone() ?? new HtmlRenderArchiveOptions();
        packaging.Validate();
        HtmlRenderOptions rendering = result.Request.ResolveOptions();
        OfficeImageExportFormat format = HtmlRenderResultImageExtensions.ResolveImageFormat(result.Request.Encoder);
        return HtmlRenderEngine.ExecuteWithDeadline(rendering, cancellationToken, token =>
            ExportArchiveCore(result, rendering, packaging, format, token));
    }

    private static HtmlRenderArchiveResult ExportArchiveCore(HtmlRenderResult result,
        HtmlRenderOptions rendering, HtmlRenderArchiveOptions packaging,
        OfficeImageExportFormat format, CancellationToken cancellationToken) {
        var images = new List<OfficeImageExportResult>(result.OutputSurfaces.Count);
        HtmlRenderResultImageExtensions.ExportImagesCore(
            result, format, rendering, images.Add, cancellationToken);

        string extension = format.GetFileExtension();
        var artifacts = new List<ArchiveArtifact>(images.Count);
        var pages = new List<HtmlRenderArchivePage>(images.Count);
        for (int index = 0; index < images.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportResult image = images[index];
            byte[] bytes = image.Bytes;
            string entryName = "pages/page-" + (index + 1).ToString("D4", System.Globalization.CultureInfo.InvariantCulture) + extension;
            string sha256 = ComputeSha256(bytes, cancellationToken);
            artifacts.Add(new ArchiveArtifact(entryName, bytes));
            pages.Add(new HtmlRenderArchivePage(
                result.OutputSurfaces[index],
                image,
                entryName,
                sha256,
                image.Diagnostics.Skip(Math.Min(result.Diagnostics.Count, image.Diagnostics.Count))));
        }

        var manifest = new HtmlRenderArchiveManifest(result, format, pages);
        byte[] manifestBytes = Encoding.UTF8.GetBytes(manifest.ToJson());
        if (manifestBytes.LongLength > packaging.MaximumManifestBytes) {
            throw new HtmlRenderArchiveLimitException(
                nameof(HtmlRenderArchiveOptions.MaximumManifestBytes),
                manifestBytes.LongLength,
                packaging.MaximumManifestBytes);
        }

        using var output = new BoundedMemoryStream(packaging.MaximumArchiveBytes);
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ArchiveArtifact artifact in artifacts) {
                cancellationToken.ThrowIfCancellationRequested();
                WriteEntry(archive, artifact.EntryName, artifact.Bytes, cancellationToken);
            }
            WriteEntry(archive, "manifest.json", manifestBytes, cancellationToken);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new HtmlRenderArchiveResult(output.ToArray(), manifest);
    }

    private static void WriteEntry(ZipArchive archive, string entryName, byte[] bytes,
        CancellationToken cancellationToken) {
        ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.NoCompression);
        entry.LastWriteTime = StableTimestamp;
        using Stream stream = entry.Open();
        const int chunkSize = 81920;
        int offset = 0;
        while (offset < bytes.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(chunkSize, bytes.Length - offset);
            stream.Write(bytes, offset, count);
            offset += count;
        }
    }

    private static string ComputeSha256(byte[] bytes, CancellationToken cancellationToken) {
        using SHA256 sha = SHA256.Create();
        const int chunkSize = 81920;
        int offset = 0;
        while (offset < bytes.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(chunkSize, bytes.Length - offset);
            sha.TransformBlock(bytes, offset, count, null, 0);
            offset += count;
        }
        sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        byte[] digest = sha.Hash ?? throw new CryptographicException("SHA-256 did not produce a digest.");
        var builder = new StringBuilder(digest.Length * 2);
        foreach (byte value in digest) builder.Append(value.ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
        return builder.ToString();
    }

    private sealed class ArchiveArtifact {
        internal ArchiveArtifact(string entryName, byte[] bytes) {
            EntryName = entryName;
            Bytes = bytes;
        }
        internal string EntryName { get; }
        internal byte[] Bytes { get; }
    }

    private sealed class BoundedMemoryStream : Stream {
        private readonly long _maximumBytes;
        private readonly MemoryStream _inner = new MemoryStream();

        internal BoundedMemoryStream(long maximumBytes) {
            _maximumBytes = maximumBytes;
        }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => true;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => _inner.Position = value; }

        internal byte[] ToArray() => _inner.ToArray();

        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);

        public override void Write(byte[] buffer, int offset, int count) {
            Ensure(count);
            _inner.Write(buffer, offset, count);
        }

#if NET8_0_OR_GREATER
        public override void Write(ReadOnlySpan<byte> buffer) {
            Ensure(buffer.Length);
            _inner.Write(buffer);
        }
#endif

        public override void WriteByte(byte value) {
            Ensure(1);
            _inner.WriteByte(value);
        }

        public override void SetLength(long value) {
            if (value > _maximumBytes) {
                throw new HtmlRenderArchiveLimitException(
                    nameof(HtmlRenderArchiveOptions.MaximumArchiveBytes), value, _maximumBytes);
            }
            _inner.SetLength(value);
        }

        private void Ensure(int count) {
            long attempted = Position > _maximumBytes - count ? long.MaxValue : Position + count;
            if (attempted > _maximumBytes) {
                throw new HtmlRenderArchiveLimitException(
                    nameof(HtmlRenderArchiveOptions.MaximumArchiveBytes), attempted, _maximumBytes);
            }
        }


        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
