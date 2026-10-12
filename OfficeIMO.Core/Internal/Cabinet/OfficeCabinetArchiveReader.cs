using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeIMO.Core.Internal.Cabinet;

/// <summary>Reads bounded, single-cabinet MS-CAB archives with uncompressed, MSZIP or LZX folders.</summary>
internal static class OfficeCabinetArchiveReader {
    /// <summary>Reads from the start of a seekable stream and restores its original position without closing it.</summary>
    internal static IReadOnlyList<OfficeCabinetEntry> Read(Stream stream, long? maxInputBytes,
        long maxExpandedBytes, long maxEntryBytes, int maxEntries, CancellationToken cancellationToken = default,
        Func<string, bool>? entryFilter = null, OfficeCabinetWorkBudget? workBudget = null) =>
        Read(stream, maxInputBytes, maxExpandedBytes, maxEntryBytes, maxEntries, out _, cancellationToken, entryFilter, workBudget);

    /// <summary>Also returns all validated decoded folder bytes, including unselected members and folder gaps.</summary>
    internal static IReadOnlyList<OfficeCabinetEntry> Read(Stream stream, long? maxInputBytes,
        long maxExpandedBytes, long maxEntryBytes, int maxEntries, out long expandedBytes,
        CancellationToken cancellationToken = default, Func<string, bool>? entryFilter = null,
        OfficeCabinetWorkBudget? workBudget = null) =>
        Read(stream, maxInputBytes, maxExpandedBytes, maxEntryBytes, maxEntries, out expandedBytes,
            cancellationToken, entryFilter, workBudget, null);

    /// <summary>Validates every CFFILE name and length before applying selection, without retaining excluded payloads.</summary>
    internal static IReadOnlyList<OfficeCabinetEntry> Read(Stream stream, long? maxInputBytes,
        long maxExpandedBytes, long maxEntryBytes, int maxEntries, out long expandedBytes,
        CancellationToken cancellationToken, Func<string, bool>? entryFilter,
        OfficeCabinetWorkBudget? workBudget, Action<string, long>? entryValidator) {
        expandedBytes = 0;
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The CAB stream must be readable.", nameof(stream));
        if (!stream.CanSeek) throw new ArgumentException("The CAB stream must be seekable.", nameof(stream));
        ValidateLimits(maxExpandedBytes, maxEntryBytes, maxEntries);
        if (maxInputBytes.HasValue && maxInputBytes.Value < 1) throw new ArgumentOutOfRangeException(nameof(maxInputBytes));
        cancellationToken.ThrowIfCancellationRequested();
        long length = stream.Length;
        if (length > int.MaxValue || (maxInputBytes.HasValue && length > maxInputBytes.Value)) {
            throw new OfficeCabinetException("CAB_INPUT_LIMIT", "The CAB archive exceeds the configured input or managed buffer limit.");
        }
        long originalPosition = stream.Position;
        try {
            stream.Position = 0;
            var data = new byte[(int)length];
            int offset = 0;
            while (offset < data.Length) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = stream.Read(data, offset, Math.Min(data.Length - offset, 65536));
                if (read == 0) throw new OfficeCabinetException("CAB_TRUNCATED", "The CAB stream ended unexpectedly.", offset);
                offset += read;
            }
            return Read(data, maxExpandedBytes, maxEntryBytes, maxEntries, out expandedBytes,
                cancellationToken, entryFilter, workBudget, entryValidator);
        } finally {
            stream.Position = originalPosition;
        }
    }

    /// <summary>Validates all members and folder data; a filter only controls which decoded members are retained.</summary>
    internal static IReadOnlyList<OfficeCabinetEntry> Read(byte[] data, long maxExpandedBytes,
        long maxEntryBytes, int maxEntries, CancellationToken cancellationToken = default,
        Func<string, bool>? entryFilter = null, OfficeCabinetWorkBudget? workBudget = null) =>
        Read(data, maxExpandedBytes, maxEntryBytes, maxEntries, out _, cancellationToken, entryFilter, workBudget);

    internal static IReadOnlyList<OfficeCabinetEntry> Read(byte[] data, long maxExpandedBytes,
        long maxEntryBytes, int maxEntries, out long expandedBytes, CancellationToken cancellationToken = default,
        Func<string, bool>? entryFilter = null, OfficeCabinetWorkBudget? workBudget = null) =>
        Read(data, maxExpandedBytes, maxEntryBytes, maxEntries, out expandedBytes,
            cancellationToken, entryFilter, workBudget, null);

    private static IReadOnlyList<OfficeCabinetEntry> Read(byte[] data, long maxExpandedBytes,
        long maxEntryBytes, int maxEntries, out long expandedBytes, CancellationToken cancellationToken,
        Func<string, bool>? entryFilter, OfficeCabinetWorkBudget? workBudget, Action<string, long>? entryValidator) {
        expandedBytes = 0;
        if (data == null) throw new ArgumentNullException(nameof(data));
        ValidateLimits(maxExpandedBytes, maxEntryBytes, maxEntries);
        cancellationToken.ThrowIfCancellationRequested();
        return new Reader(data, maxExpandedBytes, maxEntryBytes, maxEntries,
            cancellationToken, workBudget ?? new OfficeCabinetWorkBudget()).Read(entryFilter, entryValidator, out expandedBytes);
    }

    private static void ValidateLimits(long maxExpandedBytes, long maxEntryBytes, int maxEntries) {
        if (maxExpandedBytes < 0) throw new ArgumentOutOfRangeException(nameof(maxExpandedBytes));
        if (maxEntryBytes < 1) throw new ArgumentOutOfRangeException(nameof(maxEntryBytes));
        if (maxEntries < 1) throw new ArgumentOutOfRangeException(nameof(maxEntries));
    }

    private sealed class Reader {
        private readonly byte[] _data;
        private readonly long _maxExpandedBytes;
        private readonly long _maxEntryBytes;
        private readonly int _maxEntries;
        private readonly CancellationToken _cancellationToken;
        private readonly OfficeCabinetWorkBudget _workBudget;
        private int _length;

        internal Reader(byte[] data, long maxExpandedBytes, long maxEntryBytes, int maxEntries,
            CancellationToken cancellationToken, OfficeCabinetWorkBudget workBudget) {
            _data = data;
            _length = data.Length;
            _maxExpandedBytes = maxExpandedBytes;
            _maxEntryBytes = maxEntryBytes;
            _maxEntries = maxEntries;
            _cancellationToken = cancellationToken;
            _workBudget = workBudget;
        }

        internal IReadOnlyList<OfficeCabinetEntry> Read(Func<string, bool>? entryFilter,
            Action<string, long>? entryValidator, out long expandedBytes) {
            expandedBytes = 0;
            if (_length < 36 || _data[0] != (byte)'M' || _data[1] != (byte)'S' || _data[2] != (byte)'C' || _data[3] != (byte)'F') {
                throw Error("CAB_SIGNATURE", "The file is not a Microsoft Cabinet archive.");
            }
            uint declaredSize = UInt32(8);
            uint filesOffset = UInt32(16);
            int folderCount = UInt16(26);
            int fileCount = UInt16(28);
            // MS-CAB requires readers to ignore reserved CFHEADER flag bits.
            ushort flags = (ushort)(UInt16(30) & 7);
            if (UInt32(4) != 0 || UInt32(12) != 0 || UInt32(20) != 0 ||
                declaredSize < 36 || declaredSize > _length || folderCount < 1 || fileCount > _maxEntries ||
                _data[24] != 3 || _data[25] != 1) {
                throw Error("CAB_HEADER", "The CAB header contains invalid reserved fields, sizes, version or entry counts.");
            }
            // Signed cabinets may have an appended signature, but offsets must
            // remain within the size declared by CFHEADER.
            _length = (int)declaredSize;
            if ((flags & 3) != 0) throw Error("CAB_MULTIPART", "Cabinet sets and continued cabinets are not supported.");

            int offset = 36;
            int folderReserve = 0;
            int dataReserve = 0;
            if ((flags & 4) != 0) {
                Ensure(offset, 4, "CAB reserve header");
                int headerReserve = UInt16(offset);
                folderReserve = _data[offset + 2];
                dataReserve = _data[offset + 3];
                Ensure(offset, 4 + headerReserve, "CAB reserve header");
                offset += 4 + headerReserve;
            }
            Ensure(offset, folderCount * (8 + folderReserve), "CFFOLDER table");
            var folders = new FolderHeader[folderCount];
            for (int index = 0; index < folders.Length; index++) {
                _cancellationToken.ThrowIfCancellationRequested();
                folders[index] = new FolderHeader(UInt32(offset), UInt16(offset + 4), UInt16(offset + 6));
                offset += 8 + folderReserve;
            }
            if (filesOffset > int.MaxValue || filesOffset < offset) {
                throw Error("CAB_FILE_TABLE", "The CAB file table overlaps its header or is too large.");
            }
            offset = (int)filesOffset;
            Ensure(offset, fileCount * 17, "CFFILE table");
            var files = new FileHeader[fileCount];
            var folderMembers = new List<int>?[folderCount];
            var selected = new bool[fileCount];
            long extractedBytes = 0;
            for (int index = 0; index < files.Length; index++) {
                _cancellationToken.ThrowIfCancellationRequested();
                Ensure(offset, 16, "CFFILE");
                uint length = UInt32(offset);
                uint folderOffset = UInt32(offset + 4);
                ushort folderIndex = UInt16(offset + 8);
                ushort attributes = UInt16(offset + 14);
                string name = ReadCString(offset + 16, (attributes & 0x80) != 0, out int nextOffset);
                if (length > _maxEntryBytes || length > int.MaxValue) throw Error("CAB_ENTRY_LIMIT", "A CAB entry exceeds the configured size or managed buffer limit.");
                if (folderIndex >= 0xFFFD) throw Error("CAB_MULTIPART", "Continued CAB file entries are not supported.");
                if (folderIndex >= folders.Length) throw Error("CAB_FOLDER", "A CAB entry references a missing folder.");
                if (length > _maxExpandedBytes - extractedBytes) throw Error("CAB_EXPANDED_LIMIT", "The extracted CAB entries exceed the configured size limit.");
                extractedBytes += length;
                files[index] = new FileHeader(length, folderOffset, folderIndex, name);
                (folderMembers[folderIndex] ?? (folderMembers[folderIndex] = new List<int>())).Add(index);
                entryValidator?.Invoke(name, length);
                selected[index] = entryFilter == null || entryFilter(name);
                offset = nextOffset;
            }

            int fileTableEnd = offset;
            var decodedEntries = new OfficeCabinetEntry?[files.Length];
            for (int folderIndex = 0; folderIndex < folders.Length; folderIndex++) {
                _cancellationToken.ThrowIfCancellationRequested();
                FolderHeader folder = folders[folderIndex];
                if (folder.DataOffset < fileTableEnd) throw Error("CAB_DATA_OFFSET", "CAB folder data overlaps the file table.");
                byte[] source = DecompressFolder(folder, dataReserve, _maxExpandedBytes - expandedBytes);
                expandedBytes += source.LongLength;
                // Release each expanded folder after its selected members have
                // been copied instead of retaining every folder simultaneously.
                List<int>? memberIndexes = folderMembers[folderIndex];
                if (memberIndexes == null) continue;
                foreach (int index in memberIndexes) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    FileHeader file = files[index];
                    if ((long)file.FolderOffset + file.Length > source.LongLength) throw Error("CAB_ENTRY_RANGE", "A CAB entry extends past its folder.");
                    if (!selected[index]) continue;
                    var bytes = new byte[(int)file.Length];
                    if (bytes.Length > 0) Buffer.BlockCopy(source, (int)file.FolderOffset, bytes, 0, bytes.Length);
                    decodedEntries[index] = new OfficeCabinetEntry(file.Name, bytes);
                }
            }
            var entries = new List<OfficeCabinetEntry>(files.Length);
            foreach (OfficeCabinetEntry? entry in decodedEntries) if (entry != null) entries.Add(entry);
            return entries.AsReadOnly();
        }

        private byte[] DecompressFolder(FolderHeader folder, int dataReserve, long maxExpandedBytes) {
            if (folder.DataOffset > int.MaxValue) throw Error("CAB_DATA_OFFSET", "The CAB folder data offset is too large.");
            int offset = (int)folder.DataOffset;
            Ensure(offset, 0, "CAB folder data");
            ushort method = (ushort)(folder.Compression & 0x000F);
            if (method != 0 && method != 1 && method != 3) throw Error("CAB_COMPRESSION", "The CAB uses an unsupported compression method; Quantum is not supported.");
            if ((method != 3 && folder.Compression != method) || (method == 3 && (folder.Compression & 0xE0F0) != 0)) {
                throw Error("CAB_COMPRESSION", "The CAB compression parameters contain unsupported flags.");
            }
            int windowBits = (folder.Compression >> 8) & 0x1F;
            if (method == 3 && (windowBits < 15 || windowBits > 21)) throw Error("CAB_LZX_WINDOW", "The CAB uses an unsupported LZX window size.");
            var blocks = new BlockHeader[folder.BlockCount];
            long total = 0;
            for (int index = 0; index < blocks.Length; index++) {
                _cancellationToken.ThrowIfCancellationRequested();
                Ensure(offset, 8 + dataReserve, "CFDATA");
                uint checksum = UInt32(offset);
                int compressedLength = UInt16(offset + 4);
                int expandedLength = UInt16(offset + 6);
                int dataOffset = offset + 8 + dataReserve;
                Ensure(dataOffset, compressedLength, "CFDATA payload");
                if (expandedLength > maxExpandedBytes - total || total + expandedLength > int.MaxValue) {
                    throw Error("CAB_EXPANDED_LIMIT", "The expanded CAB folder exceeds the configured size or managed buffer limit.");
                }
                if (expandedLength > OfficeMsZipDecoder.WindowBytes) throw Error("CAB_BLOCK_SIZE", "A CAB frame exceeds the 32 KiB expanded size limit.");
                if (expandedLength == 0) throw Error("CAB_MULTIPART", "Split CFDATA records are not supported.");
                if (method == 0 && compressedLength != expandedLength) throw Error("CAB_UNCOMPRESSED", "An uncompressed CAB block has inconsistent sizes.");
                if (method == 1 && (compressedLength < 2 || compressedLength > OfficeMsZipDecoder.MaximumCompressedBytes ||
                    _data[dataOffset] != 0x43 || _data[dataOffset + 1] != 0x4B)) {
                    throw Error("CAB_MSZIP_HEADER", "The CAB contains an invalid MSZIP frame header or size.", dataOffset);
                }
                // Charge every validated CFDATA record, including repeated folder offsets,
                // unselected members and gaps. MSZIP replays prior history for each new frame.
                int dictionaryLength = method == 1 ? (int)Math.Min(total, OfficeMsZipDecoder.WindowBytes) : 0;
                _workBudget.ConsumeFrame(compressedLength, expandedLength, dataReserve, dictionaryLength, offset);
                if (checksum != 0) {
                    // MS-CAB section 3.1 checks payload bytes first, then the header
                    // and reserve bytes. Each region has its own remainder alignment.
                    uint actualChecksum = OfficeCabinetChecksum.Compute(_data, dataOffset, compressedLength);
                    actualChecksum = OfficeCabinetChecksum.Compute(_data, offset + 4, 4 + dataReserve, actualChecksum);
                    if (checksum != actualChecksum) throw Error("CAB_CHECKSUM", "A CAB data block has an invalid checksum.", offset);
                }
                blocks[index] = new BlockHeader(dataOffset, compressedLength, expandedLength);
                total += expandedLength;
                offset = dataOffset + compressedLength;
            }

            if (method == 3) {
                if (blocks.Length == 0) return Array.Empty<byte>();
                var chunks = new List<byte[]>(blocks.Length);
                var sizes = new List<int>(blocks.Length);
                foreach (BlockHeader block in blocks) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    var chunk = new byte[block.CompressedLength];
                    Buffer.BlockCopy(_data, block.DataOffset, chunk, 0, chunk.Length);
                    chunks.Add(chunk);
                    sizes.Add(block.ExpandedLength);
                }
                try {
                    return global::OfficeIMO.OfficeLzxDecoder.Decompress(chunks, sizes,
                        windowBits, maxExpandedBytes, _cancellationToken);
                } catch (global::OfficeIMO.OfficeLzxException exception) {
                    throw new OfficeCabinetException("CAB_" + exception.Code, exception.Message, null, exception);
                }
            }

            var output = new byte[(int)total];
            int outputOffset = 0;
            foreach (BlockHeader block in blocks) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (method == 0) {
                    Buffer.BlockCopy(_data, block.DataOffset, output, outputOffset, block.ExpandedLength);
                } else {
                    OfficeMsZipDecoder.Decompress(_data, block.DataOffset, block.CompressedLength,
                        output, outputOffset, block.ExpandedLength, _cancellationToken);
                }
                outputOffset += block.ExpandedLength;
            }
            return output;
        }

        private string ReadCString(int offset, bool utf8, out int nextOffset) {
            int end = offset;
            while (end < _length && _data[end] != 0) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (end - offset == 256) throw Error("CAB_STRING", "A CAB member name exceeds 256 encoded bytes.", offset);
                end++;
            }
            if (end >= _length) throw Error("CAB_STRING", "A CAB string is not null-terminated.", offset);
            if (end == offset) throw Error("CAB_STRING", "A CAB member name must contain between 1 and 256 encoded bytes.", offset);
            nextOffset = end + 1;
            if (utf8) {
                try { return new UTF8Encoding(false, true).GetString(_data, offset, end - offset); }
                catch (DecoderFallbackException exception) { throw new OfficeCabinetException("CAB_STRING", "A CAB member name contains invalid UTF-8.", offset, exception); }
            }
            var characters = new char[end - offset];
            for (int index = 0; index < characters.Length; index++) {
                byte value = _data[offset + index];
                if (value > 0x7F) {
                    throw Error("CAB_NAME_ENCODING", "CAB member names with non-ASCII bytes require the UTF-8 name attribute; locale-dependent name encodings are not supported.", offset);
                }
                characters[index] = (char)value;
            }
            return new string(characters);
        }

        private void Ensure(int offset, int count, string structure) {
            if (offset < 0 || count < 0 || offset > _length - count) {
                throw Error("CAB_TRUNCATED", "The CAB archive ended inside " + structure + ".", offset);
            }
        }

        private ushort UInt16(int offset) {
            Ensure(offset, 2, "integer");
            return (ushort)(_data[offset] | (_data[offset + 1] << 8));
        }

        private uint UInt32(int offset) {
            Ensure(offset, 4, "integer");
            return (uint)(_data[offset] | (_data[offset + 1] << 8) | (_data[offset + 2] << 16) | (_data[offset + 3] << 24));
        }

        private static OfficeCabinetException Error(string code, string message, long? offset = null) => new OfficeCabinetException(code, message, offset);
    }

    private readonly struct FolderHeader {
        internal FolderHeader(uint dataOffset, ushort blockCount, ushort compression) {
            DataOffset = dataOffset; BlockCount = blockCount; Compression = compression;
        }
        internal uint DataOffset { get; }
        internal ushort BlockCount { get; }
        internal ushort Compression { get; }
    }

    private readonly struct FileHeader {
        internal FileHeader(uint length, uint folderOffset, ushort folderIndex, string name) {
            Length = length; FolderOffset = folderOffset; FolderIndex = folderIndex; Name = name;
        }
        internal uint Length { get; }
        internal uint FolderOffset { get; }
        internal ushort FolderIndex { get; }
        internal string Name { get; }
    }

    private readonly struct BlockHeader {
        internal BlockHeader(int dataOffset, int compressedLength, int expandedLength) {
            DataOffset = dataOffset; CompressedLength = compressedLength; ExpandedLength = expandedLength;
        }
        internal int DataOffset { get; }
        internal int CompressedLength { get; }
        internal int ExpandedLength { get; }
    }
}
