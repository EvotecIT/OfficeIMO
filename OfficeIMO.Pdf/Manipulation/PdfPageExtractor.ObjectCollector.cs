using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    internal sealed class ObjectCollector {
        private static readonly string[] InheritablePageKeys = { "Resources", "MediaBox", "CropBox", "Rotate" };
        private readonly Dictionary<int, PdfIndirectObject> _sourceObjects;
        private readonly Dictionary<int, Dictionary<string, PdfObject>> _pageOverrides;
        private readonly List<int> _objectIds = new();
        private readonly HashSet<int> _visited = new();
    
        public ObjectCollector(Dictionary<int, PdfIndirectObject> sourceObjects, Dictionary<int, Dictionary<string, PdfObject>>? pageOverrides = null) {
            _sourceObjects = sourceObjects;
            _pageOverrides = pageOverrides ?? new Dictionary<int, Dictionary<string, PdfObject>>();
        }
    
        public IReadOnlyList<int> ObjectIds => _objectIds;
    
        public HashSet<int> PageObjectIds { get; } = new();
    
        public Dictionary<int, Dictionary<string, PdfObject>> MaterializedPageValues { get; } = new();
    
        public void CollectObjectGraph(PdfObject? value) {
            if (value is not null) {
                CollectReferences(value, isPageObject: false);
            }
        }
    
        public void CollectPage(int objectNumber) {
            if (!_sourceObjects.TryGetValue(objectNumber, out var indirect) || indirect.Value is not PdfDictionary pageDictionary) {
                throw new InvalidOperationException("PDF page object " + objectNumber.ToString(CultureInfo.InvariantCulture) + " was not found.");
            }
    
            PageObjectIds.Add(objectNumber);
            MaterializeInheritedPageValues(objectNumber, pageDictionary);
            CollectObject(objectNumber, isPageObject: true);
        }
    
        private void CollectObject(int objectNumber, bool isPageObject) {
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
            CollectReferences(indirect.Value, isPageObject, pageOverrides);
        }
    
        private void CollectReferences(PdfObject value, bool isPageObject, Dictionary<string, PdfObject>? pageOverrides = null) {
            switch (value) {
                case PdfReference reference:
                    if (reference.ObjectNumber >= 0 &&
                        _sourceObjects.TryGetValue(reference.ObjectNumber, out var referenced) &&
                        referenced.Generation != reference.Generation) {
                        throw BuildGenerationMismatchException(reference, referenced.Generation);
                    }
    
                    CollectObject(reference.ObjectNumber, isPageObject: false);
                    break;
                case PdfArray array:
                    foreach (var item in array.Items) {
                        CollectReferences(item, isPageObject: false);
                    }
    
                    break;
                case PdfDictionary dictionary:
                    foreach (var entry in dictionary.Items) {
                        if (isPageObject &&
                            (string.Equals(entry.Key, "Parent", StringComparison.Ordinal) ||
                            (pageOverrides is not null && pageOverrides.ContainsKey(entry.Key)))) {
                            continue;
                        }
    
                        CollectReferences(entry.Value, isPageObject: false);
                    }
    
                    if (isPageObject && pageOverrides is not null) {
                        foreach (var entry in pageOverrides) {
                            CollectReferences(entry.Value, isPageObject: false);
                        }
                    }
    
                    break;
                case PdfStream stream:
                    foreach (var entry in stream.Dictionary.Items) {
                        if (!string.Equals(entry.Key, "Length", StringComparison.Ordinal)) {
                            CollectReferences(entry.Value, isPageObject: false);
                        }
                    }
    
                    break;
            }
        }
    
        private void MaterializeInheritedPageValues(int pageObjectNumber, PdfDictionary pageDictionary) {
            foreach (string key in InheritablePageKeys) {
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
    }
}
