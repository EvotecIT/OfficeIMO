namespace OfficeIMO.Publisher.Internal;

/// <summary>Permutes complete existing sibling shape records, retaining every other byte and the enclosing lengths.</summary>
internal sealed class PublisherEscherOrderWriter {
    private readonly byte[] _source;
    private readonly PublisherEscherReader _reader;
    private readonly PublisherParseContext _context;
    private readonly IReadOnlyDictionary<uint, IReadOnlyList<uint>> _orders;
    private readonly Dictionary<uint, uint> _owners = new();
    private readonly HashSet<uint> _seenPages = new();
    private readonly List<PublisherNativeStreamEdit> _edits = new();

    private PublisherEscherOrderWriter(byte[] source, IReadOnlyDictionary<uint, IReadOnlyList<uint>> orders, PublisherParseContext context) {
        _source = source; _orders = orders; _context = context;
        _reader = new(new PublisherBinaryData(source, "Escher/EscherStm"), null, context);
        foreach (var page in orders) foreach (uint id in page.Value) {
            context.Token.ThrowIfCancellationRequested();
            _owners.Add(id, page.Key);
        }
    }

    internal static byte[]? Rewrite(byte[] source, IReadOnlyDictionary<uint, IReadOnlyList<uint>> orders,
        PublisherParseContext context, long maximumBytes) {
        var writer = new PublisherEscherOrderWriter(source, orders, context);
        if (!writer.Visit(0, source.Length, 0) || !writer._seenPages.SetEquals(orders.Keys)) return null;
        return PublisherNativeStreamEdit.Apply(source, writer._edits, maximumBytes, context.Token);
    }

    private bool Visit(int start, int end, int depth) {
        _context.CheckDepth(depth);
        IReadOnlyList<PublisherEscherRecord> records = _reader.Records(start, end);
        var pages = new Dictionary<uint, List<(uint Id, PublisherEscherRecord Record)>>();
        foreach (PublisherEscherRecord record in records) {
            _context.Token.ThrowIfCancellationRequested();
            if (record.Kind == 0xF004) {
                PublisherEscherRecord[] clients = _reader.Records(record.Offset, record.End).Where(child => child.Kind == 0xF011).ToArray();
                if (clients.Length > 1) return false;
                if (clients.Length == 0 || !_reader.ClientValues(clients[0]).TryGetValue(0x6801, out uint id)
                    || !_owners.TryGetValue(id, out uint pageId)) continue;
                if (!pages.TryGetValue(pageId, out var shapes)) pages.Add(pageId, shapes = new());
                shapes.Add((id, record));
            } else if ((record.Initial & 15) == 15 && record.Kind != 0xF001
                && !Visit(record.Offset, record.End, depth + 1)) return false;
        }
        foreach (var page in pages) {
            _context.Token.ThrowIfCancellationRequested();
            IReadOnlyList<uint> desired = _orders[page.Key];
            if (page.Value.Count != desired.Count || !_seenPages.Add(page.Key)) return false;
            var byId = new Dictionary<uint, PublisherEscherRecord>();
            foreach (var item in page.Value) {
                if (byId.ContainsKey(item.Id)) return false;
                byId.Add(item.Id, item.Record);
            }
            for (int index = 0; index < desired.Count; index++) {
                _context.Token.ThrowIfCancellationRequested();
                if (!byId.TryGetValue(desired[index], out PublisherEscherRecord? replacement)) return false;
                if (page.Value[index].Id == desired[index]) continue;
                byte[] bytes = new byte[checked(replacement.Length + 8)];
                Buffer.BlockCopy(_source, replacement.Offset - 8, bytes, 0, bytes.Length);
                PublisherEscherRecord slot = page.Value[index].Record;
                _edits.Add(new(slot.Offset - 8, slot.Length + 8, bytes));
            }
        }
        return true;
    }
}
