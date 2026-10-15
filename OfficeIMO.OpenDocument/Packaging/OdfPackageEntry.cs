namespace OfficeIMO.OpenDocument;

internal sealed class OdfPackageEntry {
    private byte[]? _data;
    private Func<CancellationToken, Stream>? _openRead;
    private long _declaredLength;
    private bool _decrypted;
    private XDocument? _xml;

    internal OdfPackageEntry(string name, byte[] data, string? mediaType, DateTimeOffset lastWriteTime, bool isNew) {
        Name = name;
        _data = data;
        MediaType = mediaType;
        LastWriteTime = lastWriteTime;
        IsNew = isNew;
        IsDirty = isNew;
    }

    internal OdfPackageEntry(string name, long length, Func<CancellationToken, Stream> openRead,
        DateTimeOffset lastWriteTime) {
        Name = name; _declaredLength = length; _openRead = openRead; LastWriteTime = lastWriteTime;
    }

    internal string Name { get; }
    internal string? MediaType { get; set; }
    internal DateTimeOffset LastWriteTime { get; }
    internal bool IsNew { get; private set; }
    internal bool IsDirty { get; private set; }
    internal bool IsRemoved { get; private set; }

    internal OdfPackageEntryPayloadState PayloadState => _data == null ? OdfPackageEntryPayloadState.Deferred
        : _decrypted ? OdfPackageEntryPayloadState.Decrypted : OdfPackageEntryPayloadState.Loaded;
    internal long ByteLength => IsDirty && _xml != null ? GetBytesForSave().LongLength : _data?.LongLength ?? _declaredLength;

    internal byte[] GetOriginalBytes() {
        if (_data != null) return _data;
        using Stream source = _openRead!(CancellationToken.None);
        byte[] bytes = OfficeIMO.Core.Internal.OfficeArchiveSafety.ReadEntryBytes(source, _declaredLength, _declaredLength, CancellationToken.None);
        if (bytes.LongLength != _declaredLength) throw new InvalidDataException($"OpenDocument entry '{Name}' length changed while reading.");
        _data = bytes;
        _openRead = null;
        return _data;
    }

    internal Stream OpenRead(CancellationToken token) {
        token.ThrowIfCancellationRequested();
        return _data == null ? _openRead!(token) : new MemoryStream(GetBytesForSave(), writable: false);
    }

    internal OdfPackageEntry CloneForSerialization(bool cloneXml) => new OdfPackageEntry(Name,
        _data ?? Array.Empty<byte>(), MediaType, LastWriteTime, IsNew) {
        _data = _data, _openRead = _openRead, _declaredLength = _declaredLength, _decrypted = _decrypted,
        _xml = _xml == null || !cloneXml ? _xml : new XDocument(_xml),
        IsDirty = IsDirty,
        IsRemoved = IsRemoved
    };

    internal XDocument GetXml(long maxCharacters, int maxDepth) {
        return _xml ??= OdfXmlCodec.Load(GetOriginalBytes(), Name, maxCharacters, maxDepth);
    }

    internal byte[] GetBytesForSave() {
        return IsDirty && _xml != null ? OdfXmlCodec.Save(_xml) : GetOriginalBytes();
    }

    internal void MarkDirty() {
        if (IsRemoved) throw new InvalidOperationException($"Package entry '{Name}' has been removed.");
        IsDirty = true;
    }

    internal void ReplaceBytes(byte[] data, string? mediaType) {
        _data = data ?? throw new ArgumentNullException(nameof(data));
        _openRead = null; _decrypted = false;
        _xml = null;
        MediaType = mediaType;
        IsRemoved = false;
        IsDirty = true;
    }

    internal void ReplaceLoadedBytes(byte[] data, bool decrypted = false) {
        _data = data ?? throw new ArgumentNullException(nameof(data));
        _openRead = null; _decrypted = decrypted;
        _xml = null;
        IsDirty = false;
    }

    internal void Remove() {
        IsRemoved = true;
        IsDirty = true;
    }

    internal void AcceptChanges() {
        if (IsRemoved) return;
        if (IsDirty && _xml != null) _data = OdfXmlCodec.Save(_xml);
        IsNew = false;
        IsDirty = false;
    }
}
