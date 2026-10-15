using System;
using System.Collections.Generic;

namespace OfficeIMO.Core.Internal.Cabinet;

internal static partial class OfficeCabinetArchiveReader {
    private sealed partial class Reader {
        /// <summary>Parses the canonical CAB directory once for extraction and passive inspection.</summary>
        internal Directory ReadDirectory() {
            ValidateHeader();
            uint declaredSize = UInt32(8);
            uint filesOffset = UInt32(16);
            int folderCount = UInt16(26);
            int fileCount = UInt16(28);
            // MS-CAB requires readers to ignore reserved CFHEADER flag bits.
            ushort flags = (ushort)(UInt16(30) & 7);
            if (fileCount > _maxEntries) throw Error("CAB_ENTRY_COUNT_LIMIT", "CAB file records exceed the configured entry count.");
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
                offset = nextOffset;
            }

            return new Directory(folders, files, dataReserve, offset, _length);
        }

        /// <summary>Checks intrinsic fixed-header facts before authored counts can consume a configured budget.</summary>
        internal void ValidateHeader() {
            if (_length < 36 || _data[0] != (byte)'M' || _data[1] != (byte)'S' || _data[2] != (byte)'C' || _data[3] != (byte)'F') {
                throw Error("CAB_SIGNATURE", "The file is not a Microsoft Cabinet archive.");
            }
            uint declaredSize = UInt32(8);
            if (UInt32(4) != 0 || UInt32(12) != 0 || UInt32(20) != 0 ||
                declaredSize < 36 || declaredSize > _length || UInt16(26) < 1 ||
                _data[24] != 3 || _data[25] != 1) {
                throw Error("CAB_HEADER", "The CAB header contains invalid reserved fields, sizes, version or entry counts.");
            }
        }
    }
    internal sealed class Directory {
        internal Directory(FolderHeader[] folders, FileHeader[] files, int dataReserve, int fileTableEnd, int length) {
            Folders = folders; Files = files; DataReserve = dataReserve; FileTableEnd = fileTableEnd; Length = length;
        }
        internal FolderHeader[] Folders { get; }
        internal FileHeader[] Files { get; }
        internal int DataReserve { get; }
        internal int FileTableEnd { get; }
        internal int Length { get; }
    }
}
