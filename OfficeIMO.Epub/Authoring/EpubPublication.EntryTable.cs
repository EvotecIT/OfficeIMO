using System.Collections;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    // Retained entries keyed by container path. It also counts each path's Unicode-normalized form under
    // case-insensitive comparison, so a new path can be checked for collisions without scanning every entry.
    private sealed class EntryTable : IDictionary<string, byte[]>, IReadOnlyDictionary<string, byte[]> {
        private readonly Dictionary<string, byte[]> _entries;
        private readonly Dictionary<string, int> _folded = new(StringComparer.OrdinalIgnoreCase);
        private int _unnormalizable;

        internal EntryTable(Dictionary<string, byte[]> entries) {
            _entries = entries;
            foreach (string key in entries.Keys) Track(key, 1);
        }

        public byte[] this[string key] {
            get => _entries[key];
            set {
                if (!_entries.ContainsKey(key)) Track(key, 1);
                _entries[key] = value;
            }
        }

        public Dictionary<string, byte[]>.KeyCollection Keys => _entries.Keys;
        public Dictionary<string, byte[]>.ValueCollection Values => _entries.Values;
        public int Count => _entries.Count;
        public bool ContainsKey(string key) => _entries.ContainsKey(key);
        public bool TryGetValue(string key, out byte[] value) {
            bool found = _entries.TryGetValue(key, out byte[]? retained);
            value = retained!;
            return found;
        }

        public void Add(string key, byte[] value) {
            _entries.Add(key, value);
            Track(key, 1);
        }

        public bool Remove(string key) {
            if (!_entries.Remove(key)) return false;
            Track(key, -1);
            return true;
        }

        public void Clear() {
            _entries.Clear();
            _folded.Clear();
            _unnormalizable = 0;
        }

        // True when an entry equals the path after NFC normalization and case-insensitive comparison,
        // or null when a retained path can't be normalized and the caller must compare directly.
        internal bool? HasNormalizedCollision(string path) {
            if (_unnormalizable != 0) return null;
            return _folded.ContainsKey(path.Normalize(NormalizationForm.FormC));
        }

        private void Track(string key, int change) {
            string folded;
            try { folded = key.Normalize(NormalizationForm.FormC); } catch (ArgumentException) { _unnormalizable += change; return; }
            _folded.TryGetValue(folded, out int count);
            count += change;
            if (count > 0) _folded[folded] = count; else _folded.Remove(folded);
        }

        public Dictionary<string, byte[]>.Enumerator GetEnumerator() => _entries.GetEnumerator();
        IEnumerator<KeyValuePair<string, byte[]>> IEnumerable<KeyValuePair<string, byte[]>>.GetEnumerator() => _entries.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => _entries.GetEnumerator();
        ICollection<string> IDictionary<string, byte[]>.Keys => _entries.Keys;
        ICollection<byte[]> IDictionary<string, byte[]>.Values => _entries.Values;
        IEnumerable<string> IReadOnlyDictionary<string, byte[]>.Keys => _entries.Keys;
        IEnumerable<byte[]> IReadOnlyDictionary<string, byte[]>.Values => _entries.Values;
        bool ICollection<KeyValuePair<string, byte[]>>.IsReadOnly => false;
        void ICollection<KeyValuePair<string, byte[]>>.Add(KeyValuePair<string, byte[]> item) => Add(item.Key, item.Value);
        bool ICollection<KeyValuePair<string, byte[]>>.Contains(KeyValuePair<string, byte[]> item) =>
            ((ICollection<KeyValuePair<string, byte[]>>)_entries).Contains(item);
        void ICollection<KeyValuePair<string, byte[]>>.CopyTo(KeyValuePair<string, byte[]>[] array, int arrayIndex) =>
            ((ICollection<KeyValuePair<string, byte[]>>)_entries).CopyTo(array, arrayIndex);
        bool ICollection<KeyValuePair<string, byte[]>>.Remove(KeyValuePair<string, byte[]> item) {
            if (!((ICollection<KeyValuePair<string, byte[]>>)_entries).Remove(item)) return false;
            Track(item.Key, -1);
            return true;
        }
    }
}
