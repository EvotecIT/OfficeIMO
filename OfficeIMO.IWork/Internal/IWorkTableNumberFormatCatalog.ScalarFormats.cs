namespace OfficeIMO.IWork.Internal;

internal sealed partial class IWorkTableNumberFormatCatalog {
    private readonly Dictionary<(uint Key, bool Boolean), bool> _scalarFormats = new();

    /// <summary>Accepts only the independent identity formats: text (260) and Boolean (1), with no other settings.</summary>
    internal bool IsDefaultScalarFormat(uint key, bool boolean) {
        _source.CancellationToken.ThrowIfCancellationRequested();
        if (!_initialized) Initialize();
        if (_scalarFormats.TryGetValue((key, boolean), out bool cached)) return cached;
        bool supported = false;
        if (_entries.TryGetValue(key, out var entry)) {
            string path = IWorkTableCatalogIndex.EntryPath(entry.Position) + "/6";
            byte[]? bytes = entry.Message.GetBytes(6);
            if (entry.Message.TotalFieldCount != entry.Message.FieldCount(1)
                    + entry.Message.FieldCount(2) + entry.Message.FieldCount(6)
                || entry.Message.FieldCount(2) > 1 || entry.Message.HasUnexpectedWireKind(2, IWorkWireKind.Varint)
                || entry.Message.FieldCount(6) != 1
                || entry.Message.HasUnexpectedWireKind(6, IWorkWireKind.Bytes) || bytes == null) {
                _references.Declarations.Record(_list!, path, entry.Message.FieldCount(6),
                    IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            } else {
                try {
                    IWorkWireMessage format = entry.Message.ParseNestedMessage(bytes);
                    supported = format.TotalFieldCount == 1 && format.FieldCount(1) == 1
                        && !format.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                        && format.GetUnsigned(1) == (boolean ? 1ul : 260ul);
                    if (!supported) _references.Declarations.Record(_list!, path, 1,
                        IWorkSourceDeclarationIssueKind.UnsupportedField);
                } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                    _references.Declarations.Record(_list!, path, 1);
                }
            }
        }
        _scalarFormats.Add((key, boolean), supported);
        return supported;
    }
}
