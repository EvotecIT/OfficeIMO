using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Archives;

internal static class ZipArchiveContent {
    private const string MemberLengthMismatch = "ZIP member length does not match its declaration.";

    internal static void Read(Stream source, ArchiveLimits limits, Action<string, Stream, long> consume, CancellationToken token) {
        // Scan metadata before opening payloads. Raw compressed windows avoid runtime-specific decoded-size clamps;
        // WriteStream proves decoded EOF and the declared length, and CRC binds the emitted bytes to the record.
        var records = ReadDirectory(source, limits, token);
        foreach (OfficeArchiveSafety.ZipCentralDirectoryRecord record in records) {
            ZipMemberScheme scheme = ZipEncryption.Describe(record);
            token.ThrowIfCancellationRequested();
            ZipEncryption.Validate(source, record, scheme);
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
                using (var payload = OpenMember(source, record, limits.ZipPassword, name, token)) {
                    if (payload.Read(new byte[1], 0, 1) != 0) throw new InvalidDataException("A ZIP directory member has invalid empty content: " + name);
                    payload.Finish();
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
            using var payload = OpenMember(source, item.Record, limits.ZipPassword, item.Path, token);
            consume(item.Path, payload, item.Record.UncompressedSize);
            payload.Finish();
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

    internal static ArchiveInspectionDiagnostic? Unsupported(OfficeArchiveSafety.ZipCentralDirectoryRecord record, string? password, out ZipMemberScheme scheme) {
        ArchiveInspectionDiagnostic? problem = ZipEncryption.Inspect(record, out scheme);
        if (problem == null && scheme.Encrypted && password == null)
            return new ArchiveInspectionDiagnostic(PasswordRequired, "The ZIP member is encrypted; supply ZipPassword to read it.");
        return problem;
    }

    /// <summary>Diagnostic code for encrypted members inspected without a password.</summary>
    internal const string PasswordRequired = "ZIP_PASSWORD";

    /// <summary>Opens a member's decoded content with length and checksum verification.</summary>
    internal static ZipMemberStream OpenMember(Stream source, OfficeArchiveSafety.ZipCentralDirectoryRecord record, string? password,
        string name, CancellationToken token) {
        ZipMemberScheme scheme = ZipEncryption.Describe(record);
        return new ZipMemberStream(ZipEncryption.Open(source, record, scheme, password, name, token), record, scheme, name, token);
    }

    internal static void Charge(OfficeArchiveSafety.ZipCentralDirectoryRecord record, ArchiveLimits limits) {
        limits.Operation.ReserveRead(record.UncompressedSize, record.UncompressedSize, Work(record));
    }

    internal static long Work(OfficeArchiveSafety.ZipCentralDirectoryRecord record) =>
        record.UncompressedSize > (long.MaxValue - record.CompressedSize) / 2
            ? long.MaxValue : record.CompressedSize + 2 * record.UncompressedSize;

    internal static void CopyVerified(Stream source, OfficeArchiveSafety.ZipCentralDirectoryRecord record, string? password, string name,
        Stream destination, CancellationToken token) {
        using var payload = OpenMember(source, record, password, name, token);
        var buffer = new byte[65536];
        while (true) {
            token.ThrowIfCancellationRequested();
            int read = payload.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            destination.Write(buffer, 0, read);
        }
        payload.Finish();
    }

    internal static byte[] ReadVerified(Stream source, OfficeArchiveSafety.ZipCentralDirectoryRecord record, string? password, string name,
        CancellationToken token) {
        if (record.UncompressedSize > int.MaxValue) throw ArchiveLimits.Exceeded("MaxEntryBytes", int.MaxValue);
        using var payload = OpenMember(source, record, password, name, token);
        byte[] result = payload.Guard(() => ReadDeclaredPayload(payload, (int)record.UncompressedSize, token,
            "ZIP member is shorter than its declared length."));
        payload.Finish();
        return result;
    }

    internal static byte[] ReadDeclaredPayload(Stream payload, int length, CancellationToken token,
        string shortReadMessage = MemberLengthMismatch) {
        var result = length == 0 ? Array.Empty<byte>() : new byte[length];
        for (int offset = 0; offset < result.Length;) {
            token.ThrowIfCancellationRequested();
            int read = payload.Read(result, offset, Math.Min(65536, result.Length - offset));
            if (read == 0) throw new InvalidDataException(shortReadMessage);
            offset += read;
        }
        return result;
    }

    /// <summary>Forwards a member's decoded bytes while enforcing its declared length and accumulating the ZIP (IEEE)
    /// CRC-32 shared with PNG. For traditional ZipCrypto, whose password check is a single byte, a decode failure
    /// after the check passed is reported as a wrong password or damaged member.</summary>
    internal sealed class ZipMemberStream : Stream {
        private readonly Stream _inner;
        private readonly OfficeArchiveSafety.ZipCentralDirectoryRecord _record;
        private readonly ZipMemberScheme _scheme;
        private readonly string _name;
        private readonly CancellationToken _token;
        private uint _crc = OfficePngCrc32.Begin();
        private long _read;
        internal ZipMemberStream(Stream inner, OfficeArchiveSafety.ZipCentralDirectoryRecord record, ZipMemberScheme scheme, string name,
            CancellationToken token) {
            _inner = inner; _record = record; _scheme = scheme; _name = name; _token = token;
        }

        /// <summary>Proves decoded EOF at the declared length and checks the CRC where the scheme stores one.</summary>
        internal void Finish() => Guard(() => {
            if (ReadByte() != -1 || _read != _record.UncompressedSize) throw new InvalidDataException(MemberLengthMismatch);
            if (_scheme.ChecksCrc && OfficePngCrc32.Complete(_crc) != _record.Crc32)
                throw new InvalidDataException("ZIP member checksum does not match: " + _name);
            return true;
        });

        internal T Guard<T>(Func<T> action) {
            if (_scheme.Cipher != ZipCipher.ZipCrypto) return action();
            try {
                return action();
            } catch (InvalidDataException error) {
                throw new ArchivePasswordException(_name, true, $"The password does not open ZIP member '{_name}', or the member is damaged.", error);
            }
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) {
            _token.ThrowIfCancellationRequested();
            int read = _inner.Read(buffer, offset, count);
            if (read > _record.UncompressedSize - _read) throw new InvalidDataException("ZIP member exceeds its declared length.");
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
