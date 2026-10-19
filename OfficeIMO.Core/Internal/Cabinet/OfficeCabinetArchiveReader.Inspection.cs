using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Core.Internal.Cabinet;

internal static partial class OfficeCabinetArchiveReader {
    /// <summary>Indexes an immutable bounded CAB buffer without decoding any folder.</summary>
    internal static Inspection Inspect(byte[] data, long maxExpandedBytes, long maxEntryBytes, int maxEntries,
        OfficeCabinetDirectoryBudget directoryBudget, CancellationToken token) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        ValidateLimits(maxExpandedBytes, maxEntryBytes, maxEntries);
        return new Inspection(data, maxExpandedBytes, maxEntryBytes, maxEntries, directoryBudget, token);
    }

    /// <summary>Retains the validated directory; complete member reads use the existing codecs and frame validation.</summary>
    internal sealed partial class Inspection {
        private readonly byte[] _data;
        private readonly Directory _directory;
        private readonly long _maxEntryBytes;
        private readonly int _maxEntries;
        internal Inspection(byte[] data, long maxExpandedBytes, long maxEntryBytes, int maxEntries,
            OfficeCabinetDirectoryBudget directoryBudget, CancellationToken token) {
            _data = data; _maxEntryBytes = maxEntryBytes; _maxEntries = maxEntries;
            var reader = new Reader(data, maxExpandedBytes, maxEntryBytes, maxEntries, token, new OfficeCabinetWorkBudget());
            reader.ValidateHeader();
            // Charge authored counts before allocating/parsing their tables.
            directoryBudget.Consume(reader.UInt16(26) + (long)reader.UInt16(28));
            _directory = reader.ReadDirectory();
            var compressed = new long[_directory.Folders.Length];
            var expanded = new long[_directory.Folders.Length];
            var unsupported = new string?[_directory.Folders.Length];
            var checksumFrames = new int[_directory.Folders.Length];
            var work = new long[_directory.Folders.Length];
            long total = 0;
            for (int index = 0; index < _directory.Folders.Length; index++) {
                token.ThrowIfCancellationRequested();
                FolderHeader folder = _directory.Folders[index];
                directoryBudget.Consume(folder.BlockCount);
                if (folder.DataOffset < _directory.FileTableEnd || folder.DataOffset > int.MaxValue)
                    throw Error("CAB_DATA_OFFSET", "CAB folder data overlaps the directory or has an invalid offset.");
                int offset = (int)folder.DataOffset;
                reader.Ensure(offset, 0, "CAB folder data");
                unsupported[index] = CompressionDiagnostic(folder.Compression);
                for (int block = 0; block < folder.BlockCount; block++) {
                    token.ThrowIfCancellationRequested();
                    reader.Ensure(offset, 8 + _directory.DataReserve, "CFDATA");
                    if (reader.UInt32(offset) != 0) checksumFrames[index]++;
                    int compressedLength = reader.UInt16(offset + 4);
                    int expandedLength = reader.UInt16(offset + 6);
                    int dataOffset = offset + 8 + _directory.DataReserve;
                    reader.Ensure(dataOffset, compressedLength, "CFDATA payload");
                    reader.ValidateFrame(dataOffset, compressedLength, expandedLength,
                        unsupported[index] == null ? folder.Compression & 15 : -1);
                    if (expandedLength > maxExpandedBytes - total)
                        throw Error("CAB_EXPANDED_LIMIT", "Declared CAB folder expansion exceeds the configured limit.");
                    if (expanded[index] + expandedLength > int.MaxValue)
                        throw Error("CAB_FOLDER_BUFFER_LIMIT", "Declared CAB folder expansion exceeds the managed buffer limit.");
                    int dictionary = (folder.Compression & 15) == 1 ? (int)Math.Min(expanded[index], OfficeMsZipDecoder.WindowBytes) : 0;
                    work[index] += OfficeCabinetWorkBudget.FrameWork(compressedLength, expandedLength, _directory.DataReserve, dictionary);
                    compressed[index] += compressedLength; expanded[index] += expandedLength; total += expandedLength;
                    offset += 8 + _directory.DataReserve + compressedLength;
                }
            }
            var members = new List<OfficeCabinetMemberInfo>(_directory.Files.Length);
            for (int index = 0; index < _directory.Files.Length; index++) {
                token.ThrowIfCancellationRequested();
                FileHeader file = _directory.Files[index];
                if ((long)file.FolderOffset + file.Length > expanded[file.FolderIndex])
                    throw Error("CAB_ENTRY_RANGE", "A CAB member extends beyond its declared folder expansion.");
                members.Add(new OfficeCabinetMemberInfo(index, file.Name, file.Length, file.FolderIndex,
                    _directory.Folders[file.FolderIndex].Compression, compressed[file.FolderIndex], expanded[file.FolderIndex],
                    _directory.Folders[file.FolderIndex].BlockCount, checksumFrames[file.FolderIndex],
                    work[file.FolderIndex], unsupported[file.FolderIndex]));
            }
            Members = members.AsReadOnly(); ExpandedBytes = total;
        }
        internal IReadOnlyList<OfficeCabinetMemberInfo> Members { get; }
        internal long ExpandedBytes { get; }
        internal byte[] ReadMember(int ordinal, long maxExpandedBytes, OfficeCabinetWorkBudget workBudget,
            CancellationToken token) {
            if (ordinal < 0 || ordinal >= Members.Count) throw new ArgumentOutOfRangeException(nameof(ordinal));
            return ReadFolderMembers(new[] { ordinal }, maxExpandedBytes, workBudget, token)[0];
        }
    }

    private static string? CompressionDiagnostic(ushort compression) {
        int method = compression & 15;
        if (method != 0 && method != 1 && method != 3) return "CAB_COMPRESSION";
        if (method != 3 && compression != method || method == 3 && (compression & 0xE0F0) != 0) return "CAB_COMPRESSION";
        int windowBits = (compression >> 8) & 31;
        return method == 3 && (windowBits < 15 || windowBits > 21) ? "CAB_LZX_WINDOW" : null;
    }
}
