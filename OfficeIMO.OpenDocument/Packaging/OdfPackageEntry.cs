namespace OfficeIMO.OpenDocument;

internal sealed class OdfPackageEntry {
    private byte[] _data;
    private XDocument? _xml;

    internal OdfPackageEntry(string name, byte[] data, string? mediaType, DateTimeOffset lastWriteTime, bool isNew) {
        Name = name;
        _data = data;
        MediaType = mediaType;
        LastWriteTime = lastWriteTime;
        IsNew = isNew;
        IsDirty = isNew;
    }

    internal string Name { get; }
    internal string? MediaType { get; set; }
    internal DateTimeOffset LastWriteTime { get; }
    internal bool IsNew { get; }
    internal bool IsDirty { get; private set; }
    internal bool IsRemoved { get; private set; }

    internal byte[] GetOriginalBytes() => _data;

    internal XDocument GetXml(long maxCharacters, int maxDepth) {
        return _xml ??= OdfXmlCodec.Load(_data, Name, maxCharacters, maxDepth);
    }

    internal byte[] GetBytesForSave() {
        return IsDirty && _xml != null ? OdfXmlCodec.Save(_xml) : _data;
    }

    internal void MarkDirty() {
        if (IsRemoved) throw new InvalidOperationException($"Package entry '{Name}' has been removed.");
        IsDirty = true;
    }

    internal void ReplaceBytes(byte[] data, string? mediaType) {
        _data = data ?? throw new ArgumentNullException(nameof(data));
        _xml = null;
        MediaType = mediaType;
        IsRemoved = false;
        IsDirty = true;
    }

    internal void Remove() {
        IsRemoved = true;
        IsDirty = true;
    }
}
