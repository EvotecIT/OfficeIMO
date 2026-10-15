using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Core.Internal.Cabinet;

internal static partial class OfficeCabinetArchiveReader {
    internal sealed partial class Inspection {
        /// <summary>Decodes one complete folder once, then copies the selected complete members in request order.</summary>
        internal byte[][] ReadFolderMembers(IReadOnlyList<int> ordinals, long maxExpandedBytes,
            OfficeCabinetWorkBudget workBudget, CancellationToken token) {
            if (ordinals.Count == 0) return Array.Empty<byte[]>();
            int folderIndex = -1;
            for (int index = 0; index < ordinals.Count; index++) {
                token.ThrowIfCancellationRequested();
                int ordinal = ordinals[index];
                if (ordinal < 0 || ordinal >= Members.Count) throw new ArgumentOutOfRangeException(nameof(ordinals));
                int currentFolder = _directory.Files[ordinal].FolderIndex;
                if (folderIndex >= 0 && currentFolder != folderIndex)
                    throw new ArgumentException("Selected members must belong to one CAB folder.", nameof(ordinals));
                folderIndex = currentFolder;
            }
            var reader = new Reader(_data, maxExpandedBytes, _maxEntryBytes, _maxEntries, token, workBudget);
            reader._length = _directory.Length;
            byte[] folder = reader.DecompressFolder(_directory.Folders[folderIndex], _directory.DataReserve, maxExpandedBytes);
            var result = new byte[ordinals.Count][];
            for (int index = 0; index < ordinals.Count; index++) {
                token.ThrowIfCancellationRequested();
                FileHeader file = _directory.Files[ordinals[index]];
                if ((long)file.FolderOffset + file.Length > folder.LongLength)
                    throw Error("CAB_ENTRY_RANGE", "A CAB member extends beyond its decoded folder.");
                var bytes = new byte[(int)file.Length];
                if (bytes.Length != 0) Buffer.BlockCopy(folder, (int)file.FolderOffset, bytes, 0, bytes.Length);
                result[index] = bytes;
            }
            token.ThrowIfCancellationRequested();
            return result;
        }
    }
}
