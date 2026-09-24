#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    internal sealed class ObjectCollector {
        private static readonly string[] InheritablePageKeys = { "Resources", "MediaBox", "CropBox", "Rotate" };
        private readonly Dictionary<int, PdfIndirectObject> _sourceObjects;
        private readonly Dictionary<int, Dictionary<string, PdfObject>> _pageOverrides;
        private readonly List<int> _objectIds = new();
        private readonly HashSet<int> _visited = new();
        private TraversalStack _pending;
        private readonly List<KeyValuePair<string, PdfObject>> _reverseEntries = new();
        private readonly CancellationToken _cancellationToken;
    
        public ObjectCollector(
            Dictionary<int, PdfIndirectObject> sourceObjects,
            Dictionary<int, Dictionary<string, PdfObject>>? pageOverrides = null,
            CancellationToken cancellationToken = default) {
            _sourceObjects = sourceObjects;
            _pageOverrides = pageOverrides ?? new Dictionary<int, Dictionary<string, PdfObject>>();
            _cancellationToken = cancellationToken;
        }
    
        public IReadOnlyList<int> ObjectIds => _objectIds;
    
        public HashSet<int> PageObjectIds { get; } = new();
    
        public Dictionary<int, Dictionary<string, PdfObject>> MaterializedPageValues { get; } = new();
    
        public void CollectObjectGraph(PdfObject? value) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (value is not null) {
                CollectReferences(value, isPageObject: false);
            }
        }
    
        public void CollectPage(int objectNumber) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (!_sourceObjects.TryGetValue(objectNumber, out var indirect) || indirect.Value is not PdfDictionary pageDictionary) {
                throw new InvalidOperationException("PDF page object " + objectNumber.ToString(CultureInfo.InvariantCulture) + " was not found.");
            }
    
            PageObjectIds.Add(objectNumber);
            MaterializeInheritedPageValues(objectNumber, pageDictionary);
            CollectObject(objectNumber, isPageObject: true);
        }
    
        private void CollectObject(int objectNumber, bool isPageObject) {
            QueueObject(objectNumber, isPageObject, ref _pending);
            TraversePending(ref _pending);
        }

        private void QueueObject(int objectNumber, bool isPageObject, ref TraversalStack pending) {
            if (!_visited.Add(objectNumber)) {
                return;
            }
    
            if (!_sourceObjects.TryGetValue(objectNumber, out var indirect)) {
                if (objectNumber < 0) {
                    return;
                }
    
                throw new InvalidOperationException("PDF object " + objectNumber.ToString(CultureInfo.InvariantCulture) + " was referenced but not found.");
            }
    
            _objectIds.Add(objectNumber);
            _pageOverrides.TryGetValue(objectNumber, out var pageOverrides);
            pending.Push(new TraversalItem(indirect.Value, isPageObject, pageOverrides));
        }
    
        private void CollectReferences(PdfObject value, bool isPageObject, Dictionary<string, PdfObject>? pageOverrides = null) {
            _pending.Push(new TraversalItem(value, isPageObject, pageOverrides));
            TraversePending(ref _pending);
        }

        private void TraversePending(ref TraversalStack pending) {
            try {
                while (pending.Count != 0) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    TraversalItem current = pending.Pop();
                    PdfObject value = current.Value;
                    bool isPageObject = current.IsPageObject;
                    Dictionary<string, PdfObject>? pageOverrides = current.PageOverrides;
                    switch (value) {
                    case PdfReference reference:
                        if (reference.ObjectNumber >= 0 &&
                            _sourceObjects.TryGetValue(reference.ObjectNumber, out var referenced) &&
                            referenced.Generation != reference.Generation) {
                            throw BuildGenerationMismatchException(reference, referenced.Generation);
                        }

                        QueueObject(reference.ObjectNumber, isPageObject: false, ref pending);
                        break;
                    case PdfArray array:
                        for (int index = array.Items.Count - 1; index >= 0; index--) {
                            _cancellationToken.ThrowIfCancellationRequested();
                            pending.Push(new TraversalItem(array.Items[index], isPageObject: false, pageOverrides: null));
                        }

                        break;
                    case PdfDictionary dictionary:
                        if (isPageObject && pageOverrides is not null) {
                            QueueEntriesInReverse(pageOverrides, ref pending, skipPageParent: false, pageOverrides: null, skipStreamLength: false);
                        }

                        QueueEntriesInReverse(dictionary.Items, ref pending, skipPageParent: isPageObject, pageOverrides, skipStreamLength: false);

                        break;
                    case PdfStream stream:
                        QueueEntriesInReverse(stream.Dictionary.Items, ref pending, skipPageParent: false, pageOverrides: null, skipStreamLength: true);

                        break;
                    }
                }
            } finally {
                pending.Release();
            }
        }

        private void QueueEntriesInReverse(
            Dictionary<string, PdfObject> entries,
            ref TraversalStack pending,
            bool skipPageParent,
            Dictionary<string, PdfObject>? pageOverrides,
            bool skipStreamLength) {
            try {
                foreach (KeyValuePair<string, PdfObject> entry in entries) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    _reverseEntries.Add(entry);
                }
                for (int index = _reverseEntries.Count - 1; index >= 0; index--) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    KeyValuePair<string, PdfObject> entry = _reverseEntries[index];
                    if (skipPageParent &&
                        (string.Equals(entry.Key, "Parent", StringComparison.Ordinal) ||
                        (pageOverrides is not null && pageOverrides.ContainsKey(entry.Key)))) continue;
                    if (skipStreamLength && string.Equals(entry.Key, "Length", StringComparison.Ordinal)) continue;
                    pending.Push(new TraversalItem(entry.Value, isPageObject: false, pageOverrides: null));
                }
            } finally {
                _reverseEntries.Clear();
            }
        }
    
        private void MaterializeInheritedPageValues(int pageObjectNumber, PdfDictionary pageDictionary) {
            foreach (string key in InheritablePageKeys) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (pageDictionary.Items.ContainsKey(key)) {
                    continue;
                }
    
                var inherited = ResolveInheritedValue(pageDictionary, key);
                if (inherited is null) {
                    continue;
                }
    
                if (!MaterializedPageValues.TryGetValue(pageObjectNumber, out var values)) {
                    values = new Dictionary<string, PdfObject>(StringComparer.Ordinal);
                    MaterializedPageValues[pageObjectNumber] = values;
                }
    
                values[key] = inherited;
                CollectReferences(inherited, isPageObject: false);
            }
        }
    
        private PdfObject? ResolveInheritedValue(PdfDictionary pageDictionary, string key) {
            PdfDictionary? current = pageDictionary;
            int guard = 0;
            while (current is not null && guard++ < 100) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (current.Items.TryGetValue(key, out var value)) {
                    return value;
                }
    
                if (!current.Items.TryGetValue("Parent", out var parentObj) ||
                    parentObj is not PdfReference parentReference ||
                    !PdfObjectLookup.TryGet(_sourceObjects, parentReference, out var parentIndirect) ||
                    parentIndirect.Value is not PdfDictionary parentDictionary) {
                    return null;
                }
    
                current = parentDictionary;
            }
    
            return null;
        }

        private struct TraversalStack {
#if NET8_0_OR_GREATER
            private TraversalItem[]? _items;
            private int _count;

            internal readonly int Count => _count;

            internal void Push(TraversalItem item) {
                if (_items is null) {
                    _items = ArrayPool<TraversalItem>.Shared.Rent(16);
                } else if (_count == _items.Length) {
                    TraversalItem[] expanded = ArrayPool<TraversalItem>.Shared.Rent(checked(_count * 2));
                    Array.Copy(_items, expanded, _count);
                    ArrayPool<TraversalItem>.Shared.Return(_items, clearArray: true);
                    _items = expanded;
                }

                _items[_count++] = item;
            }

            internal TraversalItem Pop() {
                if (_count == 0) throw new InvalidOperationException("The PDF object traversal stack is empty.");
                int index = --_count;
                TraversalItem item = _items![index];
                _items[index] = default;
                return item;
            }

            internal void Release() {
                if (_items is null) return;
                ArrayPool<TraversalItem>.Shared.Return(_items, clearArray: true);
                _items = null;
                _count = 0;
            }
#else
            private Stack<TraversalItem>? _items;

            internal readonly int Count => _items?.Count ?? 0;

            internal void Push(TraversalItem item) =>
                (_items ??= new Stack<TraversalItem>()).Push(item);

            internal TraversalItem Pop() => _items!.Pop();

            internal void Release() {
                // Preserve the existing retained Stack<T> behavior on legacy targets.
                if (_items is null) return;
            }
#endif
        }

        private readonly struct TraversalItem {
            internal TraversalItem(PdfObject value, bool isPageObject, Dictionary<string, PdfObject>? pageOverrides) {
                Value = value;
                IsPageObject = isPageObject;
                PageOverrides = pageOverrides;
            }

            internal PdfObject Value { get; }
            internal bool IsPageObject { get; }
            internal Dictionary<string, PdfObject>? PageOverrides { get; }
        }
    }
}
