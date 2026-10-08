using System.Collections;
using System.Threading;

namespace OfficeIMO.Latex;

// Fixed-size chunks avoid copying a growing token buffer. Public objects are
// materialized on demand, with stable identity even under concurrent inspection.
internal sealed class LatexTokenCollection : IReadOnlyList<LatexToken> {
    private const int ChunkSize = 1024;
    private readonly LatexSourceText _source;
    private readonly Chunk[] _chunks;

    private LatexTokenCollection(LatexSourceText source, Chunk[] chunks, int count) {
        _source = source;
        _chunks = chunks;
        Count = count;
        Views = new ViewList(this);
    }

    public int Count { get; }
    internal IReadOnlyList<LatexTokenView> Views { get; }

    public LatexToken this[int index] {
        get {
            ValidateIndex(index);
            Chunk chunk = _chunks[index / ChunkSize];
            int slot = index % ChunkSize;
            LatexToken?[]? cache = Volatile.Read(ref chunk.Tokens);
            if (cache == null) {
                var created = new LatexToken?[chunk.Count];
                cache = Interlocked.CompareExchange(ref chunk.Tokens, created, null) ?? created;
            }
            LatexToken? existing = Volatile.Read(ref cache[slot]);
            if (existing != null) return existing;
            // Records can be released only after every cache slot is published.
            // Holding the local array also keeps concurrent internal reads safe.
            LatexTokenRecord[]? records = Volatile.Read(ref chunk.Records);
            if (records == null) return Volatile.Read(ref cache[slot])!;
            LatexTokenRecord record = records[slot];
            var token = new LatexToken(record.Kind, _source, record.Value,
                record.StartOffset, record.EndOffset, record.IsTerminated);
            existing = Interlocked.CompareExchange(ref cache[slot], token, null);
            if (existing != null) return existing;
            if (Interlocked.Decrement(ref chunk.Remaining) == 0) Volatile.Write(ref chunk.Records, null);
            return token;
        }
    }

    public IEnumerator<LatexToken> GetEnumerator() {
        for (int index = 0; index < Count; index++) yield return this[index];
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    private LatexTokenView GetView(int index) {
        ValidateIndex(index);
        Chunk chunk = _chunks[index / ChunkSize];
        int slot = index % ChunkSize;
        LatexTokenRecord[]? records = Volatile.Read(ref chunk.Records);
        if (records != null) return new LatexTokenView(_source, records[slot]);
        LatexToken token = Volatile.Read(ref chunk.Tokens)![slot]!;
        return new LatexTokenView(_source, new LatexTokenRecord(token.Kind, token.Value,
            token.StartOffset, token.EndOffset, token.IsTerminated));
    }

    private void ValidateIndex(int index) {
        if ((uint)index >= (uint)Count) throw new ArgumentOutOfRangeException(nameof(index));
    }

    private sealed class ViewList : IReadOnlyList<LatexTokenView> {
        private readonly LatexTokenCollection _owner;
        internal ViewList(LatexTokenCollection owner) => _owner = owner;
        public int Count => _owner.Count;
        public LatexTokenView this[int index] => _owner.GetView(index);
        public IEnumerator<LatexTokenView> GetEnumerator() {
            for (int index = 0; index < Count; index++) yield return this[index];
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private sealed class Chunk {
        internal Chunk(int capacity) => Records = new LatexTokenRecord[capacity];
        internal LatexTokenRecord[]? Records;
        internal LatexToken?[]? Tokens;
        internal int Count;
        internal int Remaining;
    }

    internal sealed class Builder {
        private readonly LatexSourceText _source;
        private readonly List<Chunk> _chunks = new();
        internal Builder(LatexSourceText source) => _source = source;
        internal int Count { get; private set; }
        internal void Add(LatexTokenRecord record) {
            int slot = Count % ChunkSize;
            if (slot == 0) _chunks.Add(new Chunk(Count == 0 ? Math.Min(16, _source.Text.Length) : ChunkSize));
            Chunk chunk = _chunks[_chunks.Count - 1];
            if (slot == chunk.Records!.Length) Array.Resize(ref chunk.Records, Math.Min(ChunkSize, chunk.Records.Length * 2));
            chunk.Records![slot] = record;
            chunk.Count++;
            Count++;
        }
        internal LatexTokenCollection Build() {
            foreach (Chunk chunk in _chunks) chunk.Remaining = chunk.Count;
            return new LatexTokenCollection(_source, _chunks.ToArray(), Count);
        }
    }
}
