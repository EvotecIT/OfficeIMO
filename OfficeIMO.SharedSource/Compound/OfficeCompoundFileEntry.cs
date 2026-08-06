using System;

#if OFFICEIMO_READER_CORE
namespace OfficeIMO.Reader.Internal.Compound {
#else
namespace OfficeIMO.Core.Internal {
#endif
    /// <summary>
    /// Directory entry decoded from an OLE compound document.
    /// </summary>
    internal sealed class OfficeCompoundFileEntry {
        internal OfficeCompoundFileEntry(string name, string path, byte objectType, long size,
            bool isFallback = false, Guid classId = default, uint stateBits = 0,
            ulong creationTime = 0, ulong modifiedTime = 0, int directoryOrder = int.MaxValue) {
            Name = name;
            Path = path;
            ObjectType = objectType;
            Size = size;
            IsFallback = isFallback;
            ClassId = classId;
            StateBits = stateBits;
            CreationTime = creationTime;
            ModifiedTime = modifiedTime;
            DirectoryOrder = directoryOrder;
        }

        internal string Name { get; }

        internal string Path { get; }

        internal byte ObjectType { get; }

        internal long Size { get; }

        internal Guid ClassId { get; }

        internal uint StateBits { get; }

        internal ulong CreationTime { get; }

        internal ulong ModifiedTime { get; }

        /// <summary>Depth-first directory traversal order from the source compound file.</summary>
        internal int DirectoryOrder { get; }

        /// <summary>True when the entry is a synthetic unqualified lookup for an otherwise unreachable directory item.</summary>
        internal bool IsFallback { get; }

        internal bool IsStorage => ObjectType == 1 || ObjectType == 5;

        internal bool IsStream => ObjectType == 2;
    }
}
