using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Archives;

internal static class ZipArchiveContent {
    private const ushort ZipStored = 0;
    private const ushort ZipDeflate = 8;

    internal static void Read(Stream source, ArchiveLimits limits, Action<string, Stream, long> consume, CancellationToken token) {
        // Scan metadata before opening payloads. Raw compressed windows avoid runtime-specific decoded-size clamps;
        // WriteStream proves decoded EOF and the declared length, and CRC binds the emitted bytes to the record.
        var records = ReadDirectory(source, limits, token);
        foreach (OfficeArchiveSafety.ZipCentralDirectoryRecord record in records) {
            if (record.CompressionMethod != ZipStored && record.CompressionMethod != ZipDeflate)
                throw new NotSupportedException("ZIP members must be stored or Deflate-compressed.");
            if (!OfficeArchiveSafety.HasSupportedZipFlags(record.Flags, record.CompressionMethod))
                throw new NotSupportedException("ZIP member uses encrypted, patched, enhanced or reserved compression flags.");
            token.ThrowIfCancellationRequested();
            OfficeArchiveSafety.ValidateEntryPayload(source, record);
        }
        var paths = new ArchivePaths(limits);
        var selected = new List<(OfficeArchiveSafety.ZipCentralDirectoryRecord Record, string Path)>();
        long declared = 0;
        foreach (OfficeArchiveSafety.ZipCentralDirectoryRecord record in records) {
            token.ThrowIfCancellationRequested();
            long length = record.UncompressedSize;
            string name = ArchiveZipNames.Decode(record, limits.ZipLegacyCodePage);
            if (name.EndsWith("/", StringComparison.Ordinal) || name.EndsWith("\\", StringComparison.Ordinal)) {
                if (length != 0) throw new InvalidDataException("A ZIP directory member has content: " + name);
                Charge(record, limits);
                using (var payload = new Crc32ReadStream(OfficeArchiveSafety.OpenEntryPayload(source, record, token), record.UncompressedSize, token)) {
                    if (payload.Read(new byte[1], 0, 1) != 0 || payload.Crc32 != record.Crc32)
                        throw new InvalidDataException("A ZIP directory member has invalid empty content: " + name);
                }
                paths.AddDirectory(name.Substring(0, name.Length - 1));
                continue;
            }
            string path = paths.Add(name);
            limits.CheckEntry(length);
            if (!limits.Select(path)) continue;
            // Declared lengths are binding: WriteStream never writes more than them and rejects short members.
            ArchiveLimits.Add(ref declared, length, limits.Expanded, nameof(ArchiveExtractionOptions.MaxExpandedBytes));
            selected.Add((record, path));
        }
        foreach (var item in selected) {
            token.ThrowIfCancellationRequested();
            Charge(item.Record, limits);
            using var payload = new Crc32ReadStream(OfficeArchiveSafety.OpenEntryPayload(source, item.Record, token), item.Record.UncompressedSize, token);
            consume(item.Path, payload, item.Record.UncompressedSize);
            payload.VerifyComplete();
            if (payload.Crc32 != item.Record.Crc32) throw new InvalidDataException("ZIP member checksum does not match: " + item.Path);
        }
    }

    internal static List<OfficeArchiveSafety.ZipCentralDirectoryRecord> ReadDirectory(Stream source, ArchiveLimits limits, CancellationToken token) {
        var records = new List<OfficeArchiveSafety.ZipCentralDirectoryRecord>();
        source.Position = 0;
        OfficeArchiveSafety.ZipCentralDirectoryScanResult scan =
            OfficeArchiveSafety.ScanZipCentralDirectoryForExtraction(source, source.Length, limits.Entries, records, token);
        if (scan.LimitExceeded) throw ArchiveLimits.Exceeded(nameof(ArchiveExtractionOptions.MaxEntryCount), limits.Entries);
        if (!scan.IsValid) throw new InvalidDataException(scan.Error ?? "The ZIP central directory is malformed.");
        limits.Operation.AdmitEntries(records.Count);
        return records;
    }

    internal static ArchiveInspectionDiagnostic? Unsupported(OfficeArchiveSafety.ZipCentralDirectoryRecord record) {
        if (record.CompressionMethod != ZipStored && record.CompressionMethod != ZipDeflate)
            return new ArchiveInspectionDiagnostic("ZIP_COMPRESSION", "The ZIP member compression method is unsupported.");
        if (!OfficeArchiveSafety.HasSupportedZipFlags(record.Flags, record.CompressionMethod))
            return new ArchiveInspectionDiagnostic("ZIP_FLAGS", "The ZIP member uses encryption or unsupported flags.");
        return null;
    }

    internal static void Charge(OfficeArchiveSafety.ZipCentralDirectoryRecord record, ArchiveLimits limits) {
        limits.Operation.ReadZip(record.UncompressedSize, record.UncompressedSize, Work(record));
    }

    internal static long Work(OfficeArchiveSafety.ZipCentralDirectoryRecord record) =>
        record.UncompressedSize > (long.MaxValue - record.CompressedSize) / 2
            ? long.MaxValue : record.CompressedSize + 2 * record.UncompressedSize;

    internal static void CopyVerified(Stream source, OfficeArchiveSafety.ZipCentralDirectoryRecord record,
        Stream destination, CancellationToken token) {
        using var payload = new Crc32ReadStream(OfficeArchiveSafety.OpenEntryPayload(source, record, token), record.UncompressedSize, token);
        var buffer = new byte[65536];
        while (true) {
            token.ThrowIfCancellationRequested();
            int read = payload.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            destination.Write(buffer, 0, read);
        }
        payload.VerifyComplete();
        if (payload.Crc32 != record.Crc32) throw new InvalidDataException("ZIP member checksum does not match.");
    }

    internal static byte[] ReadVerified(Stream source, OfficeArchiveSafety.ZipCentralDirectoryRecord record, CancellationToken token) {
        if (record.UncompressedSize > int.MaxValue) throw ArchiveLimits.Exceeded("MaxEntryBytes", int.MaxValue);
        using var payload = new Crc32ReadStream(OfficeArchiveSafety.OpenEntryPayload(source, record, token), record.UncompressedSize, token);
        var result = new byte[(int)record.UncompressedSize];
        for (int offset = 0; offset < result.Length;) {
            token.ThrowIfCancellationRequested();
            int read = payload.Read(result, offset, Math.Min(65536, result.Length - offset));
            if (read == 0) throw new InvalidDataException("ZIP member is shorter than its declared length.");
            offset += read;
        }
        payload.VerifyComplete();
        if (payload.Crc32 != record.Crc32) throw new InvalidDataException("ZIP member checksum does not match.");
        return result;
    }

    /// <summary>Forwards reads while accumulating the ZIP (IEEE) CRC-32 shared with PNG.</summary>
    private sealed class Crc32ReadStream : Stream {
        private readonly Stream _inner;
        private uint _crc = OfficePngCrc32.Begin();
        private readonly long _declared;
        private readonly CancellationToken _token;
        private long _read;
        internal Crc32ReadStream(Stream inner, long declared, CancellationToken token) { _inner = inner; _declared = declared; _token = token; }
        internal void VerifyComplete() {
            if (ReadByte() != -1 || _read != _declared) throw new InvalidDataException("ZIP member length does not match its declaration.");
        }
        internal uint Crc32 => OfficePngCrc32.Complete(_crc);
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) {
            _token.ThrowIfCancellationRequested();
            int read = _inner.Read(buffer, offset, count);
            if (read > _declared - _read) throw new InvalidDataException("ZIP member exceeds its declared length.");
            _read += read;
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
