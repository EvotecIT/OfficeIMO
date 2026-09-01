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
            var pending = new Stack<TraversalItem>();
            QueueObject(objectNumber, isPageObject, pending);
            TraversePending(pending);
        }

        private void QueueObject(int objectNumber, bool isPageObject, Stack<TraversalItem> pending) {
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
            var pending = new Stack<TraversalItem>();
            pending.Push(new TraversalItem(value, isPageObject, pageOverrides));
            TraversePending(pending);
        }

        private void TraversePending(Stack<TraversalItem> pending) {
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
    
                    QueueObject(reference.ObjectNumber, isPageObject: false, pending);
                    break;
                case PdfArray array:
                    for (int index = array.Items.Count - 1; index >= 0; index--) {
                        pending.Push(new TraversalItem(array.Items[index], isPageObject: false, pageOverrides: null));
                    }
    
                    break;
                case PdfDictionary dictionary:
                    if (isPageObject && pageOverrides is not null) {
                        foreach (var entry in pageOverrides.Reverse()) {
                            pending.Push(new TraversalItem(entry.Value, isPageObject: false, pageOverrides: null));
                        }
                    }

                    foreach (var entry in dictionary.Items.Reverse()) {
                        if (isPageObject &&
                            (string.Equals(entry.Key, "Parent", StringComparison.Ordinal) ||
                            (pageOverrides is not null && pageOverrides.ContainsKey(entry.Key)))) {
                            continue;
                        }
    
                        pending.Push(new TraversalItem(entry.Value, isPageObject: false, pageOverrides: null));
                    }
    
                    break;
                case PdfStream stream:
                    foreach (var entry in stream.Dictionary.Items.Reverse()) {
                        if (!string.Equals(entry.Key, "Length", StringComparison.Ordinal)) {
                            pending.Push(new TraversalItem(entry.Value, isPageObject: false, pageOverrides: null));
                        }
                    }
    
                    break;
                }
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
