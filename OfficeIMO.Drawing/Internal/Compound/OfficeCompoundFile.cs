using System.Collections.Generic;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>
    /// Represents an OLE compound document container with decoded stream bytes and directory entries.
    /// </summary>
    internal sealed class OfficeCompoundFile {
        internal OfficeCompoundFile(
            IReadOnlyDictionary<string, byte[]> streams,
            IReadOnlyList<OfficeCompoundFileEntry> entries) {
            Streams = streams;
            Entries = entries;
        }

        internal IReadOnlyDictionary<string, byte[]> Streams { get; }

        internal IReadOnlyList<OfficeCompoundFileEntry> Entries { get; }
    }
}
