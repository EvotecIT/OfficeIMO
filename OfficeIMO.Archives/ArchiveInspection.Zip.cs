using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives {

    public sealed partial class ArchiveInspection {
        private Stream? _zipSource;
        private List<OfficeArchiveSafety.ZipCentralDirectoryRecord>? _zipRecords;
        private long _zipWorkBytes;

        private IReadOnlyList<ArchiveInspectionEntry> ReadZipInventory(Stream input, CancellationToken token) {
            _zipSource = input;
            _zipRecords = ZipArchiveContent.ReadDirectory(_zipSource, _limits, token);
            var entries = new List<ArchiveInspectionEntry>(_zipRecords.Count);
            long total = 0;
            foreach (var record in _zipRecords) {
                token.ThrowIfCancellationRequested();
                string name = ArchiveZipNames.Decode(record, _limits.ZipLegacyCodePage);
                _limits.CheckEntry(record.UncompressedSize);
                ArchiveLimits.Add(ref total, record.UncompressedSize, _limits.Expanded, "MaxExpandedBytes");
                bool directory = name.EndsWith("/", StringComparison.Ordinal) || name.EndsWith("\\", StringComparison.Ordinal);
                var diagnostic = ZipArchiveContent.Unsupported(record);
                var compression = record.CompressionMethod == 0 ? ArchiveCompression.Store
                    : record.CompressionMethod == 8 ? ArchiveCompression.Deflate : ArchiveCompression.Unknown;
                entries.Add(new ArchiveInspectionEntry(entries.Count, name, null, name, record.UncompressedSize,
                    null, null, null, compression, record.CompressionMethod, null, null, null, null, null,
                    diagnostic, directory, record.CompressedSize, ZipTime(record.DosDate, record.DosTime)));
            }
            return entries.AsReadOnly();
        }

        private Stream OpenZipMember(int ordinal, CancellationToken token) {
            var record = _zipRecords![ordinal];
            ReserveZipRead(record);
            byte[] bytes = ZipArchiveContent.ReadVerified(_zipSource!, record, token);
            if (Entries[ordinal].IsDirectory && bytes.Length != 0) throw new InvalidDataException("ZIP directory member has content.");
            return new MemoryStream(bytes, writable: false);
        }

        private void ReserveZipRead(OfficeArchiveSafety.ZipCentralDirectoryRecord record) {
            ArchiveLimits.Check(record.UncompressedSize, _limits.Expanded - _readBytes, "MaxExpandedBytes");
            ZipArchiveContent.Charge(record, _limits);
            ArchiveLimits.Add(ref _readBytes, record.UncompressedSize, _limits.Expanded, "MaxExpandedBytes");
            _zipWorkBytes = checked(_zipWorkBytes + ZipArchiveContent.Work(record));
        }

        private void VerifyZipMember(int ordinal, Stream destination, CancellationToken token) {
            var record = _zipRecords![ordinal];
            ReserveZipRead(record);
            if (Entries[ordinal].IsDirectory && record.UncompressedSize != 0) throw new InvalidDataException("ZIP directory member has content.");
            ZipArchiveContent.CopyVerified(_zipSource!, record, destination, token);
        }

        private ArchiveReadBatch ReadZipMembers(IReadOnlyList<int> ordinals, long maximum, CancellationToken token) {
            if (maximum < 0) throw new ArgumentOutOfRangeException("maxSelectedBytes");
            if (ordinals.Count > Entries.Count) throw new ArgumentException("The request exceeds the inventory count.", nameof(ordinals));
            var selected = new int[ordinals.Count];
            var seen = new HashSet<int>();
            long bytes = 0;
            long work = 0;
            for (int index = 0; index < selected.Length; index++) {
                token.ThrowIfCancellationRequested();
                int ordinal = ordinals[index];
                if (ordinal < 0 || ordinal >= Entries.Count) throw new ArgumentOutOfRangeException(nameof(ordinals));
                if (!seen.Add(ordinal)) throw new ArgumentException("A source ordinal is repeated.", nameof(ordinals));
                var entry = Entries[ordinal];
                if (entry.Diagnostic != null) throw new NotSupportedException(entry.Diagnostic.Code + ": " + entry.Diagnostic.Message);
                ArchiveLimits.Add(ref bytes, entry.Length, maximum, "maxSelectedBytes");
                ArchiveLimits.Add(ref work, ZipArchiveContent.Work(_zipRecords![ordinal]), long.MaxValue, "Operation.MaxDecompressionWorkBytes");
                selected[index] = ordinal;
            }
            ArchiveLimits.Check(bytes, _limits.Expanded - _readBytes, "MaxExpandedBytes");
            _limits.Operation.CheckRead(bytes, bytes, work);
            long before = _readBytes;
            long workBefore = _zipWorkBytes;
            var members = new List<ArchiveReadMember>(selected.Length);
            try {
                foreach (int ordinal in selected) {
                    token.ThrowIfCancellationRequested();
                    members.Add(new ArchiveReadMember(Entries[ordinal], OpenZipMember(ordinal, token)));
                }
                token.ThrowIfCancellationRequested();
                return new ArchiveReadBatch(members, bytes, _readBytes - before, _zipWorkBytes - workBefore);
            } catch { foreach (var member in members) member.Content.Dispose(); throw; }
        }

        private static DateTimeOffset? ZipTime(ushort date, ushort time) {
            try { return new DateTimeOffset(1980 + (date >> 9), (date >> 5) & 15, date & 31,
                time >> 11, (time >> 5) & 63, (time & 31) * 2, TimeSpan.Zero); }
            catch (ArgumentOutOfRangeException) { return null; }
        }
    }
}
