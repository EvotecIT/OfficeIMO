using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Reader;

internal static partial class OfficeDocumentModelTraversal {
    /// <summary>Indexes occurrence locations independently of table cells; each match is consumed once.</summary>
    private sealed class TableMatchIndex<T> {
        private readonly Dictionary<string, Bucket> _payloads = new(StringComparer.Ordinal);

        internal void Add(string payload, ReaderLocation? location, T value) {
            if (!_payloads.TryGetValue(payload, out var bucket)) _payloads.Add(payload, bucket = new());
            bucket.Add(LocationCoordinates(location), value);
        }

        internal bool TryTake(string payload, ReaderLocation? location, out T value) {
            if (_payloads.TryGetValue(payload, out var bucket)) return bucket.TryTake(LocationCoordinates(location), out value);
            value = default!;
            return false;
        }

        internal IEnumerable<T> Candidates(string payload, ReaderLocation? location) =>
            _payloads.TryGetValue(payload, out var bucket) ? bucket.Candidates(LocationCoordinates(location)) : Array.Empty<T>();

        private sealed class Bucket {
            private readonly Dictionary<int, LocationGroup> _groups = new();
            private int _next;

            internal void Add(string?[] coordinates, T value) {
                int mask = Mask(coordinates);
                if (!_groups.TryGetValue(mask, out var group)) _groups.Add(mask, group = new(mask));
                group.Add(_next++, coordinates, value);
            }

            internal bool TryTake(string?[] coordinates, out T value) {
                int mask = Mask(coordinates);
                int first = int.MaxValue;
                LocationGroup? selected = null;
                foreach (var group in _groups.Values) {
                    int? candidate = group.First(mask, coordinates);
                    if (candidate.HasValue && candidate.Value < first) {
                        first = candidate.Value;
                        selected = group;
                    }
                }
                if (selected == null) { value = default!; return false; }
                value = selected.Take(first);
                return true;
            }

            internal IEnumerable<T> Candidates(string?[] coordinates) {
                int mask = Mask(coordinates);
                return _groups.Values.SelectMany(group => group.Candidates(mask, coordinates))
                    .OrderBy(candidate => candidate.Id).Select(candidate => candidate.Value);
            }
        }

        // Tables with the same set of known coordinates share a tuple index. A query uses only
        // coordinates known on both sides. Cache each such projection once per group, rather than
        // scanning all occurrences when several individual coordinates have large posting sets.
        private sealed class LocationGroup {
            private readonly int _mask;
            private readonly Dictionary<int, (string?[] Coordinates, T Value)> _entries = new();
            private readonly Dictionary<int, Dictionary<string, Queue<int>>> _projections = new();

            internal LocationGroup(int mask) => _mask = mask;

            internal void Add(int id, string?[] coordinates, T value) {
                _entries.Add(id, (coordinates, value));
                foreach (var projection in _projections) AddKey(projection.Value, Key(projection.Key, coordinates), id);
            }

            internal int? First(int queryMask, string?[] coordinates) {
                var ids = MatchingIds(queryMask, coordinates);
                if (ids == null) return null;
                while (ids.Count > 0 && !_entries.ContainsKey(ids.Peek())) ids.Dequeue();
                return ids.Count > 0 ? ids.Peek() : null;
            }

            internal IEnumerable<(int Id, T Value)> Candidates(int queryMask, string?[] coordinates) {
                var ids = MatchingIds(queryMask, coordinates);
                if (ids == null) yield break;
                foreach (int id in ids)
                    if (_entries.TryGetValue(id, out var entry)) yield return (id, entry.Value);
            }

            private Queue<int>? MatchingIds(int queryMask, string?[] coordinates) {
                if (_entries.Count == 0) return null;
                int shared = queryMask & _mask;
                if (!_projections.TryGetValue(shared, out var index)) {
                    index = new(StringComparer.Ordinal);
                    // IDs encode source order; dictionary enumeration order is not a framework contract.
                    foreach (var entry in _entries.OrderBy(entry => entry.Key))
                        AddKey(index, Key(shared, entry.Value.Coordinates), entry.Key);
                    _projections.Add(shared, index);
                }
                return index.TryGetValue(Key(shared, coordinates), out var ids) ? ids : null;
            }

            internal T Take(int id) {
                T value = _entries[id].Value;
                _entries.Remove(id);
                return value;
            }

            private static void AddKey(Dictionary<string, Queue<int>> index, string key, int id) {
                if (!index.TryGetValue(key, out var ids)) index.Add(key, ids = new());
                ids.Enqueue(id);
            }
        }

        private static int Mask(string?[] coordinates) {
            int mask = 0;
            for (int field = 0; field < coordinates.Length; field++)
                if (coordinates[field] != null) mask |= 1 << field;
            return mask;
        }

        private static string Key(int mask, string?[] coordinates) {
            var builder = new System.Text.StringBuilder();
            for (int field = 0; field < coordinates.Length; field++)
                if ((mask & (1 << field)) != 0) AppendIdentity(builder, coordinates[field]);
            return builder.ToString();
        }
    }

    private static string?[] LocationCoordinates(ReaderLocation? location) => new[] {
        Coordinate(location?.Path), Coordinate(location?.Sheet), Coordinate(location?.A1Range),
        Coordinate(location?.HeadingPath), Coordinate(location?.HierarchyHeadingPath), Coordinate(location?.HeadingSlug),
        Coordinate(location?.SourceBlockKind), Coordinate(location?.BlockAnchor),
        location?.Page?.ToString(CultureInfo.InvariantCulture), location?.Slide?.ToString(CultureInfo.InvariantCulture),
        location?.BlockIndex?.ToString(CultureInfo.InvariantCulture), location?.SourceBlockIndex?.ToString(CultureInfo.InvariantCulture),
        location?.StartLine?.ToString(CultureInfo.InvariantCulture), location?.EndLine?.ToString(CultureInfo.InvariantCulture),
        location?.NormalizedStartLine?.ToString(CultureInfo.InvariantCulture), location?.NormalizedEndLine?.ToString(CultureInfo.InvariantCulture),
        location?.TableIndex?.ToString(CultureInfo.InvariantCulture)
    };

    private static string? Coordinate(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;
}
