using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Core.Internal.Cabinet;

internal static partial class OfficeCabinetArchiveReader {
    internal sealed partial class Inspection {
        /// <summary>Decodes one complete folder by frames, retaining only selected complete members in request order.</summary>
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
            var files = new FileHeader[ordinals.Count];
            for (int index = 0; index < files.Length; index++) files[index] = _directory.Files[ordinals[index]];
            byte[][] result = reader.ReadFolderMembers(_directory.Folders[folderIndex], _directory.DataReserve, files, maxExpandedBytes);
            token.ThrowIfCancellationRequested();
            return result;
        }
    }
}
