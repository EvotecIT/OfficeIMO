using System.Collections;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Tracks workbook custom document properties and marks the owning document dirty when callers mutate the collection directly.
    /// </summary>
    public sealed class ExcelCustomDocumentPropertyCollection : IDictionary<string, ExcelCustomProperty>, IReadOnlyDictionary<string, ExcelCustomProperty> {
        private readonly Dictionary<string, ExcelCustomProperty> _properties;
        private Action? _changed;
        private bool _suppressChangeTracking;

        internal ExcelCustomDocumentPropertyCollection() {
            _properties = new Dictionary<string, ExcelCustomProperty>(StringComparer.OrdinalIgnoreCase);
        }

        internal void SetChangeHandler(Action changed) {
            _changed = changed ?? throw new ArgumentNullException(nameof(changed));
        }

        internal void ReplaceWith(IEnumerable<KeyValuePair<string, ExcelCustomProperty>> properties) {
            _suppressChangeTracking = true;
            try {
                DetachAll();
                _properties.Clear();
                foreach (var property in properties) {
                    string key = NormalizeKey(property.Key);
                    ExcelCustomProperty value = property.Value ?? throw new ArgumentNullException(nameof(properties));
                    Attach(value);
                    _properties[key] = value;
                }
            } finally {
                _suppressChangeTracking = false;
            }
        }

        /// <inheritdoc />
        public ExcelCustomProperty this[string key] {
            get => _properties[NormalizeKey(key)];
            set {
                string normalizedKey = NormalizeKey(key);
                if (_properties.TryGetValue(normalizedKey, out ExcelCustomProperty? previous)) {
                    Detach(previous);
                }

                Attach(value ?? throw new ArgumentNullException(nameof(value)));
                _properties[normalizedKey] = value;
                MarkChanged();
            }
        }

        /// <inheritdoc />
        public ICollection<string> Keys => _properties.Keys;

        /// <inheritdoc />
        public ICollection<ExcelCustomProperty> Values => _properties.Values;

        IEnumerable<string> IReadOnlyDictionary<string, ExcelCustomProperty>.Keys => _properties.Keys;

        IEnumerable<ExcelCustomProperty> IReadOnlyDictionary<string, ExcelCustomProperty>.Values => _properties.Values;

        /// <inheritdoc />
        public int Count => _properties.Count;

        /// <inheritdoc />
        public bool IsReadOnly => false;

        /// <inheritdoc />
        public void Add(string key, ExcelCustomProperty value) {
            string normalizedKey = NormalizeKey(key);
            Attach(value ?? throw new ArgumentNullException(nameof(value)));
            _properties.Add(normalizedKey, value);
            MarkChanged();
        }

        /// <inheritdoc />
        public bool ContainsKey(string key) {
            return _properties.ContainsKey(NormalizeKey(key));
        }

        /// <inheritdoc />
        public bool Remove(string key) {
            string normalizedKey = NormalizeKey(key);
            if (!_properties.TryGetValue(normalizedKey, out ExcelCustomProperty? property)) {
                return false;
            }

            bool removed = _properties.Remove(normalizedKey);
            if (removed) {
                Detach(property);
                MarkChanged();
            }

            return removed;
        }

        /// <inheritdoc />
        public bool TryGetValue(string key, out ExcelCustomProperty value) {
            return _properties.TryGetValue(NormalizeKey(key), out value!);
        }

        /// <inheritdoc />
        public void Add(KeyValuePair<string, ExcelCustomProperty> item) {
            Add(item.Key, item.Value);
        }

        /// <inheritdoc />
        public void Clear() {
            if (_properties.Count == 0) {
                return;
            }

            DetachAll();
            _properties.Clear();
            MarkChanged();
        }

        /// <inheritdoc />
        public bool Contains(KeyValuePair<string, ExcelCustomProperty> item) {
            return ((ICollection<KeyValuePair<string, ExcelCustomProperty>>)_properties)
                .Contains(new KeyValuePair<string, ExcelCustomProperty>(NormalizeKey(item.Key), item.Value));
        }

        /// <inheritdoc />
        public void CopyTo(KeyValuePair<string, ExcelCustomProperty>[] array, int arrayIndex) {
            ((ICollection<KeyValuePair<string, ExcelCustomProperty>>)_properties).CopyTo(array, arrayIndex);
        }

        /// <inheritdoc />
        public bool Remove(KeyValuePair<string, ExcelCustomProperty> item) {
            bool removed = ((ICollection<KeyValuePair<string, ExcelCustomProperty>>)_properties)
                .Remove(new KeyValuePair<string, ExcelCustomProperty>(NormalizeKey(item.Key), item.Value));
            if (removed) {
                Detach(item.Value);
                MarkChanged();
            }

            return removed;
        }

        /// <inheritdoc />
        public IEnumerator<KeyValuePair<string, ExcelCustomProperty>> GetEnumerator() {
            return _properties.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator() {
            return GetEnumerator();
        }

        private static string NormalizeKey(string key) {
            if (string.IsNullOrWhiteSpace(key)) {
                throw new ArgumentException("Custom property name is required.", nameof(key));
            }

            return key.Trim();
        }

        private void MarkChanged() {
            if (!_suppressChangeTracking) {
                _changed?.Invoke();
            }
        }

        private void Attach(ExcelCustomProperty property) {
            property.SetChangeHandler(MarkChanged);
        }

        private static void Detach(ExcelCustomProperty property) {
            property.SetChangeHandler(null);
        }

        private void DetachAll() {
            foreach (ExcelCustomProperty property in _properties.Values) {
                Detach(property);
            }
        }
    }
}
