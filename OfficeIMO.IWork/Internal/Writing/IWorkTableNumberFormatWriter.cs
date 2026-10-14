namespace OfficeIMO.IWork.Internal;

/// <summary>Deduplicates the explicit numeric formats selected by one owned native table.</summary>
internal sealed class IWorkTableNumberFormatWriter {
    private readonly Dictionary<(IWorkNumberFormatKind Kind, int Decimals, bool Grouping, IWorkNegativeNumberStyle Negative),
        (uint Key, uint References, IWorkNumberFormat Format)> _entries = new();

    internal IWorkTableNumberFormatWriter(IEnumerable<IWorkTableCell> cells, System.Threading.CancellationToken cancellationToken) {
        foreach (IWorkTableCell cell in cells) {
            cancellationToken.ThrowIfCancellationRequested();
            if (cell.NumberFormat is not { } format) continue;
            format.ValidateNativeNumericCreation();
            var key = Key(format);
            if (_entries.TryGetValue(key, out var entry))
                _entries[key] = (entry.Key, entry.References + 1, entry.Format);
            else _entries.Add(key, ((uint)_entries.Count + 1, 1, format));
        }
    }

    internal uint? GetKey(IWorkTableCell? cell) => cell?.NumberFormat is { } format ? _entries[Key(format)].Key : null;

    internal IWorkProtoWriter Catalog(Func<IWorkProtoWriter> create) {
        var catalog = create().UInt(1, 2).UInt(2, (uint)_entries.Count + 1);
        foreach (var entry in _entries.Values.OrderBy(entry => entry.Key))
            catalog.Message(3, create().UInt(1, entry.Key).UInt(2, entry.References)
                .Message(6, FormatFields(entry.Format, create())));
        return catalog;
    }

    internal static IWorkProtoWriter FormatFields(IWorkNumberFormat format, IWorkProtoWriter message) =>
        message.UInt(1, format.Kind switch { IWorkNumberFormatKind.Number => 256u,
                IWorkNumberFormatKind.Percentage => 258u, IWorkNumberFormatKind.Scientific => 259u,
                _ => throw new InvalidDataException("Unsupported native numeric creation format.") })
            .UInt(2, (uint)format.DecimalPlaces!.Value).UInt(4, (uint)format.NegativeStyle).Bool(5, format.ThousandsSeparator);

    private static (IWorkNumberFormatKind, int, bool, IWorkNegativeNumberStyle) Key(IWorkNumberFormat format) =>
        (format.Kind, format.DecimalPlaces!.Value, format.ThousandsSeparator, format.NegativeStyle);
}
