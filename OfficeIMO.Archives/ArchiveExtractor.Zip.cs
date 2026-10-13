using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Archives;

public static partial class ArchiveExtractor {
    // General-purpose flag bits: traditional encryption, strong encryption and masked local headers.
    private const ushort ZipEncryptionFlags = 0x0001 | 0x0040 | 0x2000;
    private const ushort ZipUtf8NameFlag = 0x0800;
    private const ushort ZipStored = 0;
    private const ushort ZipDeflate = 8;
    private static readonly Encoding StrictUtf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

    private static void ExtractZip(Stream source, string staging, ArchiveLimits limits, List<ExtractedArchiveFile> files,
        ref long total, CancellationToken token) {
        // The central directory is scanned before ZipArchive materializes it, so member-count limits, multi-disk
        // packages, encryption and ambiguous directory layouts are rejected from bounded reads. Each entry is then
        // bound to its record by exact name and sizes, so flags, method and CRC describe the member actually read.
        var records = new List<OfficeArchiveSafety.ZipCentralDirectoryRecord>();
        source.Position = 0;
        OfficeArchiveSafety.ZipCentralDirectoryScanResult scan =
            OfficeArchiveSafety.ScanZipCentralDirectoryForExtraction(source, source.Length, limits.Entries, records, token);
        if (!scan.IsValid) throw new InvalidDataException(scan.Error ?? "The ZIP central directory is malformed.");
        if (scan.LimitExceeded) throw new InvalidDataException("ZIP exceeds the configured member count.");
        foreach (OfficeArchiveSafety.ZipCentralDirectoryRecord record in records) {
            if ((record.Flags & ZipEncryptionFlags) != 0) throw new NotSupportedException("Encrypted ZIP members are not supported.");
            if (record.CompressionMethod != ZipStored && record.CompressionMethod != ZipDeflate)
                throw new NotSupportedException("ZIP members must be stored or Deflate-compressed.");
        }
        source.Position = 0;
        using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
        ReadOnlyCollection<ZipArchiveEntry> entries = archive.Entries;
        if (entries.Count != records.Count) throw new InvalidDataException("The ZIP central directory does not match its readable members.");
        var paths = new ArchivePaths(limits);
        var selected = new List<(ZipArchiveEntry Entry, string Path, uint Crc32)>();
        long declared = 0;
        for (int index = 0; index < entries.Count; index++) {
            token.ThrowIfCancellationRequested();
            ZipArchiveEntry entry = entries[index];
            OfficeArchiveSafety.ZipCentralDirectoryRecord record = records[index];
            long length = entry.Length;
            if (record.UncompressedSize != uint.MaxValue && record.UncompressedSize != length
                || record.CompressedSize != uint.MaxValue && record.CompressedSize != entry.CompressedLength)
                throw new InvalidDataException("The ZIP central directory does not match its readable members.");
            string name = entry.FullName;
            if (!string.Equals(name, DecodeName(record), StringComparison.Ordinal))
                throw new InvalidDataException("The ZIP central directory does not match its readable members.");
            if (name.EndsWith("/", StringComparison.Ordinal) || name.EndsWith("\\", StringComparison.Ordinal)) {
                if (length != 0) throw new InvalidDataException("A ZIP directory member has content: " + name);
                paths.AddDirectory(name.Substring(0, name.Length - 1));
                continue;
            }
            string path = paths.Add(name);
            limits.CheckEntry(length);
            if (!limits.Select(path)) continue;
            // Declared lengths are binding: WriteStream never writes more than them and rejects short members.
            ArchiveLimits.Add(ref declared, length, limits.Expanded, "Selected ZIP content");
            selected.Add((entry, path, record.Crc32));
        }
        foreach (var item in selected) {
            token.ThrowIfCancellationRequested();
            using var payload = new Crc32ReadStream(item.Entry.Open());
            WriteStream(staging, item.Path, payload, item.Entry.Length, limits, files, ref total, token);
            if (payload.Crc32 != item.Crc32) throw new InvalidDataException("ZIP member checksum does not match: " + item.Path);
        }
    }

    // Unflagged names decode through a host- or runtime-dependent code page, so they are accepted only as ASCII.
    // Flagged names must be well-formed UTF-8 rather than silently gaining replacement characters.
    private static string DecodeName(OfficeArchiveSafety.ZipCentralDirectoryRecord record) {
        byte[] bytes = record.NameBytes;
        if ((record.Flags & ZipUtf8NameFlag) != 0) {
            try {
                return StrictUtf8.GetString(bytes);
            } catch (DecoderFallbackException exception) {
                throw new InvalidDataException("A ZIP member name flagged as UTF-8 is not valid UTF-8.", exception);
            }
        }
        foreach (byte value in bytes)
            if (value > 0x7f) throw new NotSupportedException("ZIP member names with non-ASCII characters must be flagged as UTF-8.");
        return Encoding.ASCII.GetString(bytes);
    }

    /// <summary>Forwards reads while accumulating the ZIP (IEEE) CRC-32 shared with PNG.</summary>
    private sealed class Crc32ReadStream : Stream {
        private readonly Stream _inner;
        private uint _crc = OfficePngCrc32.Begin();
        internal Crc32ReadStream(Stream inner) => _inner = inner;
        internal uint Crc32 => OfficePngCrc32.Complete(_crc);
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = _inner.Read(buffer, offset, count);
            if (read > 0) _crc = OfficePngCrc32.Append(_crc, buffer, offset, read);
            return read;
        }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
