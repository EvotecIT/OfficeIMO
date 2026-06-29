namespace OfficeIMO.Shared {
    /// <summary>
    /// Directory entry decoded from an OLE compound document.
    /// </summary>
    internal sealed class OfficeCompoundFileEntry {
        internal OfficeCompoundFileEntry(string name, string path, byte objectType, long size) {
            Name = name;
            Path = path;
            ObjectType = objectType;
            Size = size;
        }

        internal string Name { get; }

        internal string Path { get; }

        internal byte ObjectType { get; }

        internal long Size { get; }

        internal bool IsStorage => ObjectType == 1 || ObjectType == 5;

        internal bool IsStream => ObjectType == 2;
    }
}
