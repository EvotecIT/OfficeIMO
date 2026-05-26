namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a parsed PDF document with access to pages, catalog and metadata.
/// Note: MVP reader supports classic xref tables and simple stream parsing sufficient for OfficeIMO.Pdf output.
/// </summary>
public sealed class PdfReadDocument {
    private readonly Dictionary<int, PdfIndirectObject> _objects;
    private readonly string _trailerRaw;
    private readonly PdfReadOptions _options;
    private readonly Dictionary<string, PdfNamedDestination> _nameDestinations = new(StringComparer.Ordinal);
    private readonly Dictionary<string, PdfNamedDestination> _stringDestinations = new(StringComparer.Ordinal);

    private PdfReadDocument(Dictionary<int, PdfIndirectObject> objects, string trailerRaw, PdfReadOptions? options) {
        _objects = objects; _trailerRaw = trailerRaw; _options = options ?? new PdfReadOptions();
        Pages = CollectPages();
        Metadata = ExtractMetadata();
        PageLabels = ExtractPageLabels();
        NamedDestinations = ExtractNamedDestinations();
        Outlines = ExtractOutlines();
        OpenAction = ExtractOpenAction();
        ViewerPreferences = ExtractViewerPreferences();
        FormFields = ExtractFormFields();
        CatalogPageMode = ExtractCatalogName("PageMode");
        CatalogPageLayout = ExtractCatalogName("PageLayout");
        CatalogVersion = ExtractCatalogName("Version");
        CatalogLanguage = ExtractCatalogString("Lang");
    }

    /// <summary>All page objects discovered in document order.</summary>
    public IReadOnlyList<PdfReadPage> Pages { get; }

    /// <summary>Document metadata (when present).</summary>
    public PdfMetadata Metadata { get; }

    /// <summary>Top-level document outline/bookmark entries.</summary>
    public IReadOnlyList<PdfOutlineItem> Outlines { get; }

    /// <summary>Page-label rules discovered from the document catalog.</summary>
    public IReadOnlyList<PdfPageLabel> PageLabels { get; }

    /// <summary>Named destinations discovered from the document catalog.</summary>
    public IReadOnlyList<PdfNamedDestination> NamedDestinations { get; }

    /// <summary>Simple document open action discovered from the document catalog, when supported.</summary>
    public PdfDocumentOpenAction? OpenAction { get; }

    /// <summary>Simple viewer preference entries discovered from the document catalog, when supported.</summary>
    public PdfViewerPreferences? ViewerPreferences { get; }

    /// <summary>Simple AcroForm fields discovered from the document catalog.</summary>
    public IReadOnlyList<PdfFormField> FormFields { get; }

    /// <summary>Catalog page mode, for example UseOutlines or FullScreen, when present.</summary>
    public string? CatalogPageMode { get; }

    /// <summary>Catalog page layout, for example SinglePage or TwoColumnLeft, when present.</summary>
    public string? CatalogPageLayout { get; }

    /// <summary>Catalog PDF version override, for example 1.7, when present.</summary>
    public string? CatalogVersion { get; }

    /// <summary>Catalog language tag, for example en-US or pl-PL, when present.</summary>
    public string? CatalogLanguage { get; }

    /// <summary>Loads a PDF from bytes into a typed object model.</summary>
    public static PdfReadDocument Load(byte[] pdf, PdfReadOptions? options = null) {
        var (map, trailer) = PdfSyntax.ParseObjects(pdf);
        return new PdfReadDocument(map, trailer, options);
    }

    /// <summary>Loads a PDF from a file path.</summary>
    public static PdfReadDocument Load(string path, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Load(File.ReadAllBytes(path), options);
    }

    /// <summary>Loads a PDF from the current position of a readable stream.</summary>
    public static PdfReadDocument Load(Stream stream, PdfReadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Load(buffer.ToArray(), options);
    }

    /// <summary>Extracts full‑document plain text (pages separated by blank lines).</summary>
    public string ExtractText() {
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < Pages.Count; i++) {
            if (i > 0) sb.AppendLine();
            sb.Append(Pages[i].ExtractText());
        }
        return sb.ToString();
    }

    /// <summary>Extracts image XObjects from all pages in page order.</summary>
    public IReadOnlyList<PdfExtractedImage> ExtractImages() => PdfImageExtractor.ExtractImages(this);

    private IReadOnlyList<PdfOutlineItem> ExtractOutlines() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue("Outlines", out var outlinesObj) ||
            ResolveDict(outlinesObj) is not PdfDictionary outlines ||
            !outlines.Items.TryGetValue("First", out var firstObj)) {
            return Array.Empty<PdfOutlineItem>();
        }

        var visited = new HashSet<int>();
        return ReadOutlineSiblings(firstObj, 1, visited).AsReadOnly();
    }

    private List<PdfOutlineItem> ReadOutlineSiblings(PdfObject firstObj, int level, HashSet<int> visited) {
        var items = new List<PdfOutlineItem>();
        PdfObject? currentObj = firstObj;

        while (currentObj is not null && ResolveDict(currentObj) is PdfDictionary current) {
            int objectNumber = currentObj is PdfReference reference ? reference.ObjectNumber : FindObjectNumberFor(current);
            if (objectNumber > 0 && !visited.Add(objectNumber)) {
                break;
            }

            string title = current.Get<PdfStringObj>("Title")?.Value ?? string.Empty;
            var (pageNumber, destinationTop) = GetOutlineDestination(current);
            var children = current.Items.TryGetValue("First", out var childObj)
                ? ReadOutlineSiblings(childObj, level + 1, visited)
                : new List<PdfOutlineItem>();

            items.Add(new PdfOutlineItem(title, level, pageNumber, destinationTop, children.AsReadOnly()));

            currentObj = current.Items.TryGetValue("Next", out var nextObj) ? nextObj : null;
        }

        return items;
    }

    private IReadOnlyList<PdfPageLabel> ExtractPageLabels() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue("PageLabels", out var pageLabelsObject) ||
            ResolveObject(pageLabelsObject) is not PdfDictionary tree ||
            tree.Items.ContainsKey("Kids") ||
            !tree.Items.TryGetValue("Nums", out var numsObject) ||
            ResolveObject(numsObject) is not PdfArray nums ||
            nums.Items.Count % 2 != 0) {
            return Array.Empty<PdfPageLabel>();
        }

        var labels = new List<PdfPageLabel>();
        for (int i = 0; i < nums.Items.Count; i += 2) {
            if (ResolveObject(nums.Items[i]) is not PdfNumber pageIndexNumber ||
                !TryGetNonNegativeInteger(pageIndexNumber, out int pageIndex) ||
                ResolveObject(nums.Items[i + 1]) is not PdfDictionary labelDictionary) {
                return Array.Empty<PdfPageLabel>();
            }

            string? style = null;
            if (ResolveObject(labelDictionary.Items.TryGetValue("S", out var styleObject) ? styleObject : null) is PdfName styleName &&
                !string.IsNullOrEmpty(styleName.Name)) {
                style = styleName.Name;
            }

            string? prefix = null;
            if (ResolveObject(labelDictionary.Items.TryGetValue("P", out var prefixObject) ? prefixObject : null) is PdfStringObj prefixText) {
                prefix = prefixText.Value;
            }

            int? startNumber = null;
            if (ResolveObject(labelDictionary.Items.TryGetValue("St", out var startObject) ? startObject : null) is PdfNumber startNumberObject &&
                TryGetPositiveInteger(startNumberObject, out int parsedStartNumber)) {
                startNumber = parsedStartNumber;
            }

            labels.Add(new PdfPageLabel(pageIndex, style, prefix, startNumber));
        }

        labels.Sort((left, right) => left.StartPageIndex.CompareTo(right.StartPageIndex));
        return labels.Count == 0 ? Array.Empty<PdfPageLabel>() : labels.AsReadOnly();
    }

    private (int? PageNumber, double? DestinationTop) GetOutlineDestination(PdfDictionary item) {
        if (item.Items.TryGetValue("Dest", out var destObj) &&
            TryReadDestinationOrNamedDestination(destObj, out int? pageNumber, out double? destinationTop)) {
            return (pageNumber, destinationTop);
        }

        if (item.Items.TryGetValue("A", out var actionObject) &&
            ResolveObject(actionObject) is PdfDictionary action &&
            action.Get<PdfName>("S")?.Name == "GoTo" &&
            action.Items.TryGetValue("D", out var actionDestination) &&
            TryReadDestinationOrNamedDestination(actionDestination, out pageNumber, out destinationTop)) {
            return (pageNumber, destinationTop);
        }

        return (null, null);
    }

    private IReadOnlyList<PdfNamedDestination> ExtractNamedDestinations() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null) {
            return Array.Empty<PdfNamedDestination>();
        }

        var result = new List<PdfNamedDestination>();
        if (catalog.Items.TryGetValue("Dests", out var directDests) &&
            ResolveDict(directDests) is PdfDictionary directDestinations) {
            foreach (var entry in directDestinations.Items) {
                if (TryCreateNamedDestination(entry.Key, entry.Value, out var destination)) {
                    AddNamedDestination(result, destination, PdfNamedDestinationTokenKind.Name);
                }
            }
        }

        if (catalog.Items.TryGetValue("Names", out var namesObject) &&
            ResolveDict(namesObject) is PdfDictionary namesDictionary &&
            namesDictionary.Items.TryGetValue("Dests", out var namedDestinationTree)) {
            AddNamedDestinationsFromNameTree(namedDestinationTree, result, new HashSet<int>());
        }

        return result.Count == 0 ? Array.Empty<PdfNamedDestination>() : result.AsReadOnly();
    }

    private void AddNamedDestinationsFromNameTree(
        PdfObject treeObject,
        List<PdfNamedDestination> result,
        HashSet<int> visitedReferences) {
        if (treeObject is PdfReference reference) {
            if (!visitedReferences.Add(reference.ObjectNumber) ||
                !_objects.TryGetValue(reference.ObjectNumber, out var indirect)) {
                return;
            }

            AddNamedDestinationsFromNameTree(indirect.Value, result, visitedReferences);
            return;
        }

        if (treeObject is not PdfDictionary tree) {
            return;
        }

        if (tree.Items.TryGetValue("Names", out var destinationNamesObject) &&
            ResolveArray(destinationNamesObject) is PdfArray destinationNames) {
            for (int i = 0; i + 1 < destinationNames.Items.Count; i += 2) {
                if (TryReadDestinationName(destinationNames.Items[i], out string? name, out _) &&
                    TryCreateNamedDestination(name!, destinationNames.Items[i + 1], out var destination)) {
                    AddNamedDestination(result, destination, PdfNamedDestinationTokenKind.String);
                }
            }
        }

        if (tree.Items.TryGetValue("Kids", out var kidsObject) &&
            ResolveArray(kidsObject) is PdfArray kids) {
            foreach (var kid in kids.Items) {
                AddNamedDestinationsFromNameTree(kid, result, visitedReferences);
            }
        }
    }

    private PdfDocumentOpenAction? ExtractOpenAction() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue("OpenAction", out var openActionObject)) {
            return null;
        }

        PdfObject? resolved = ResolveObject(openActionObject);
        if (resolved is PdfArray &&
            TryReadDestination(resolved, out int? pageNumber, out double? destinationTop)) {
            return new PdfDocumentOpenAction("Destination", pageNumber, destinationTop);
        }

        if (resolved is PdfDictionary dictionary &&
            dictionary.Get<PdfName>("S")?.Name == "GoTo" &&
            dictionary.Items.TryGetValue("D", out var destination) &&
            TryReadDestination(destination, out pageNumber, out destinationTop)) {
            return new PdfDocumentOpenAction("GoTo", pageNumber, destinationTop);
        }

        return null;
    }

    private PdfViewerPreferences? ExtractViewerPreferences() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue("ViewerPreferences", out var viewerPreferencesObject) ||
            ResolveObject(viewerPreferencesObject) is not PdfDictionary dictionary) {
            return null;
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var entry in dictionary.Items) {
            if (!TryFormatSimpleValue(entry.Value, out string? value)) {
                return null;
            }

            values[entry.Key] = value!;
        }

        return values.Count == 0 ? null : new PdfViewerPreferences(values);
    }

    private IReadOnlyList<PdfFormField> ExtractFormFields() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue("AcroForm", out var acroFormObject) ||
            ResolveObject(acroFormObject) is not PdfDictionary acroForm ||
            !acroForm.Items.TryGetValue("Fields", out var fieldsObject) ||
            ResolveArray(fieldsObject) is not PdfArray fields) {
            return Array.Empty<PdfFormField>();
        }

        var result = new List<PdfFormField>();
        var visited = new HashSet<int>();
        for (int i = 0; i < fields.Items.Count; i++) {
            ReadFormField(fields.Items[i], null, null, result, visited);
        }

        return result.Count == 0 ? Array.Empty<PdfFormField>() : result.AsReadOnly();
    }

    private void ReadFormField(PdfObject fieldObject, string? parentName, string? inheritedFieldType, List<PdfFormField> result, HashSet<int> visited) {
        PdfObject? resolved = ResolveObject(fieldObject);
        if (resolved is not PdfDictionary field) {
            return;
        }

        int? objectNumber = null;
        if (fieldObject is PdfReference reference) {
            objectNumber = reference.ObjectNumber;
            if (!visited.Add(reference.ObjectNumber)) {
                return;
            }
        } else {
            int foundObjectNumber = FindExactObjectNumberFor(field);
            if (foundObjectNumber > 0) {
                objectNumber = foundObjectNumber;
                if (!visited.Add(foundObjectNumber)) {
                    return;
                }
            }
        }

        string? partialName = TryReadText(field, "T");
        string? fullName = CombineFieldName(parentName, partialName);
        string? fieldType = TryReadName(field, "FT") ?? inheritedFieldType;
        string? value = TryReadSimpleFieldValue(field, "V");
        string? alternateName = TryReadText(field, "TU");
        string? mappingName = TryReadText(field, "TM");
        int? flags = TryReadInteger(field, "Ff");

        PdfArray? kids = field.Items.TryGetValue("Kids", out var kidsObject) ? ResolveArray(kidsObject) : null;
        bool hasReadableFieldState = fieldType != null || value != null || flags.HasValue;
        bool hasTerminalShape = kids is null || hasReadableFieldState;
        if (hasTerminalShape && (fullName != null || hasReadableFieldState || alternateName != null || mappingName != null)) {
            result.Add(new PdfFormField(objectNumber, fullName, partialName, fieldType, value, alternateName, mappingName, flags));
        }

        if (kids is null) {
            return;
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            ReadFormField(kids.Items[i], fullName, fieldType, result, visited);
        }
    }

    private static string? CombineFieldName(string? parentName, string? partialName) {
        if (string.IsNullOrEmpty(parentName)) {
            return string.IsNullOrEmpty(partialName) ? null : partialName;
        }

        if (string.IsNullOrEmpty(partialName)) {
            return parentName;
        }

        return parentName + "." + partialName;
    }

    private string? TryReadText(PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out var value) && ResolveObject(value) is PdfStringObj text && !string.IsNullOrEmpty(text.Value)
            ? text.Value
            : null;
    }

    private string? TryReadName(PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out var value) && ResolveObject(value) is PdfName name && !string.IsNullOrEmpty(name.Name)
            ? name.Name
            : null;
    }

    private string? TryReadSimpleFieldValue(PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out var value) || !TryFormatSimpleValue(value, out string? text)) {
            return null;
        }

        return text;
    }

    private int? TryReadInteger(PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out var value) ||
            ResolveObject(value) is not PdfNumber number ||
            number.Value < int.MinValue ||
            number.Value > int.MaxValue ||
            Math.Truncate(number.Value) != number.Value) {
            return null;
        }

        return (int)number.Value;
    }

    private void AddNamedDestination(List<PdfNamedDestination> result, PdfNamedDestination destination, PdfNamedDestinationTokenKind kind) {
        result.Add(destination);
        var lookup = kind == PdfNamedDestinationTokenKind.String ? _stringDestinations : _nameDestinations;
#if NETSTANDARD2_0
        if (!lookup.ContainsKey(destination.Name)) {
            lookup[destination.Name] = destination;
        }
#else
        lookup.TryAdd(destination.Name, destination);
#endif
    }

    private bool TryReadDestinationName(PdfObject obj, out string? name, out PdfNamedDestinationTokenKind kind) {
        PdfObject? resolved = ResolveObject(obj);
        if (resolved is PdfDictionary dictionary &&
            dictionary.Items.TryGetValue("D", out var explicitDestination)) {
            resolved = ResolveObject(explicitDestination);
        }

        switch (resolved) {
            case PdfStringObj text:
                name = text.Value;
                kind = PdfNamedDestinationTokenKind.String;
                return !string.IsNullOrEmpty(name);
            case PdfName pdfName:
                name = pdfName.Name;
                kind = PdfNamedDestinationTokenKind.Name;
                return !string.IsNullOrEmpty(name);
            default:
                name = null;
                kind = PdfNamedDestinationTokenKind.None;
                return false;
        }
    }

    private bool TryCreateNamedDestination(string name, PdfObject destinationObject, out PdfNamedDestination destination) {
        destination = null!;
        if (string.IsNullOrEmpty(name) || !TryReadDestination(destinationObject, out int? pageNumber, out double? destinationTop)) {
            return false;
        }

        destination = new PdfNamedDestination(name, pageNumber, destinationTop);
        return true;
    }

    private bool TryReadDestinationOrNamedDestination(PdfObject destinationObject, out int? pageNumber, out double? destinationTop) {
        if (TryReadDestination(destinationObject, out pageNumber, out destinationTop)) {
            return true;
        }

        if (TryReadDestinationName(destinationObject, out string? name, out var kind)) {
            var lookup = kind == PdfNamedDestinationTokenKind.String ? _stringDestinations : _nameDestinations;
            if (lookup.TryGetValue(name!, out var destination)) {
                pageNumber = destination.PageNumber;
                destinationTop = destination.DestinationTop;
                return true;
            }
        }

        pageNumber = null;
        destinationTop = null;
        return false;
    }

    private bool TryReadDestination(PdfObject destinationObject, out int? pageNumber, out double? destinationTop) {
        pageNumber = null;
        destinationTop = null;

        PdfObject? resolved = ResolveObject(destinationObject);
        if (resolved is PdfDictionary dictionary &&
            dictionary.Items.TryGetValue("D", out var explicitDestination)) {
            resolved = ResolveObject(explicitDestination);
        }

        if (resolved is not PdfArray destination || destination.Items.Count == 0) {
            return false;
        }

        if (destination.Items[0] is PdfReference pageRef) {
            pageNumber = GetPageNumberForObject(pageRef.ObjectNumber);
        }

        if (destination.Items.Count > 3 && ResolveObject(destination.Items[3]) is PdfNumber top) {
            destinationTop = top.Value;
        }

        return true;
    }

    private enum PdfNamedDestinationTokenKind {
        None,
        Name,
        String
    }

    private int? GetPageNumberForObject(int objectNumber) {
        for (int i = 0; i < Pages.Count; i++) {
            if (Pages[i].ObjectNumber == objectNumber) {
                return i + 1;
            }
        }

        return null;
    }

    private PdfDictionary? FindCatalog() {
        return PdfSyntax.FindCatalog(_objects, _trailerRaw);
    }

    private string? ExtractCatalogName(string key) {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue(key, out var value) ||
            ResolveObject(value) is not PdfName name ||
            string.IsNullOrEmpty(name.Name)) {
            return null;
        }

        return name.Name;
    }

    private string? ExtractCatalogString(string key) {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue(key, out var value) ||
            ResolveObject(value) is not PdfStringObj text ||
            string.IsNullOrEmpty(text.Value)) {
            return null;
        }

        return text.Value;
    }

    private static bool TryGetNonNegativeInteger(PdfNumber number, out int value) {
        value = 0;
        if (number.Value < 0 || number.Value > int.MaxValue || Math.Truncate(number.Value) != number.Value) {
            return false;
        }

        value = (int)number.Value;
        return true;
    }

    private static bool TryGetPositiveInteger(PdfNumber number, out int value) {
        if (TryGetNonNegativeInteger(number, out value) && value > 0) {
            return true;
        }

        value = 0;
        return false;
    }

    private bool TryFormatSimpleValue(PdfObject value, out string? text) {
        switch (ResolveObject(value)) {
            case PdfNumber number:
                text = number.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                return true;
            case PdfBoolean boolean:
                text = boolean.Value ? "true" : "false";
                return true;
            case PdfName name:
                text = name.Name;
                return true;
            case PdfStringObj stringObj:
                text = stringObj.Value;
                return true;
            case PdfNull:
                text = "null";
                return true;
            case PdfArray array:
                var parts = new List<string>(array.Items.Count);
                for (int i = 0; i < array.Items.Count; i++) {
                    if (!TryFormatSimpleValue(array.Items[i], out string? itemText)) {
                        text = null;
                        return false;
                    }

                    parts.Add(itemText!);
                }

                text = "[" + string.Join(" ", parts) + "]";
                return true;
            default:
                text = null;
                return false;
        }
    }

    private List<PdfReadPage> CollectPages() {
        // Prefer true page tree traversal when possible (Catalog -> Pages -> Kids ...)
        var result = new List<PdfReadPage>();
        PdfDictionary? catalog = FindCatalog();
        if (catalog is not null) {
            var pagesNode = ResolveDict(catalog.Items.TryGetValue("Pages", out var v) ? v : null);
            if (pagesNode is not null) {
                var kids = ResolveArray(pagesNode.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null);
                int kidCount = kids?.Items.Count ?? 0;
                var visitedNodes = new HashSet<PdfDictionary>();
                var visitedPages = new HashSet<int>();
                int? target = null;
                var cnt = pagesNode.Get<PdfNumber>("Count");
                if (cnt is not null) {
                    int cc = (int)cnt.Value; if (cc > 0) target = cc;
                }
                TraversePagesNodeDeepLimited(pagesNode, visitedNodes, visitedPages, result, target);
                if (result.Count == 0 && kidCount > 0) {
                    // Build a reachable candidate set from Kids only
                    var reachable = CollectReachableLeafCandidates(pagesNode);
                    foreach (var id in reachable) {
                        if (_objects.TryGetValue(id, out var ind) && ind.Value is PdfDictionary dict) {
                            result.Add(new PdfReadPage(id, dict, _objects));
                            if (target.HasValue && result.Count >= target.Value) break;
                        }
                        if (target.HasValue && result.Count >= target.Value) break;
                    }
                }
            }
        }
        if (result.Count > 0) return result;

        // Fallback: scan all dictionaries; accept leaf candidates whose Parent chain leads to a /Pages node
        foreach (var kv in _objects) {
            if (kv.Value.Value is PdfDictionary dict) {
                if (IsLeafPageByParent(dict)) result.Add(new PdfReadPage(kv.Key, dict, _objects));
            }
        }
        result.Sort((a, b) => a.ObjectNumber.CompareTo(b.ObjectNumber));
        return result;
    }

    private PdfDictionary? ResolveDict(PdfObject? obj) {
        if (obj is null) return null;
        if (obj is PdfDictionary d) return d;
        if (obj is PdfReference r && _objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfDictionary dd) return dd;
        return null;
    }

    private PdfObject? ResolveObject(PdfObject? obj) {
        if (obj is PdfReference reference &&
            _objects.TryGetValue(reference.ObjectNumber, out var indirect)) {
            return indirect.Value;
        }

        return obj;
    }

    private PdfArray? ResolveArray(PdfObject? obj) {
        if (obj is null) return null;
        if (obj is PdfArray a) return a;
        if (obj is PdfReference r && _objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfArray aa) return aa;
        return null;
    }

    private void TraversePagesNode(PdfDictionary node, List<PdfReadPage> outList) {
        var type = node.Get<PdfName>("Type")?.Name;
        if (type == "Page" || (type is null && IsLikelyPage(node))) {
            // Find this node's object number
            int objNum = FindObjectNumberFor(node);
            outList.Add(new PdfReadPage(objNum, node, _objects));
            return;
        }
        var kidsObj = node.Items.TryGetValue("Kids", out var kidsValue) ? kidsValue : null;
        if (type == "Pages" || (type is null && ResolveArray(kidsObj) is not null)) {
            var kids = ResolveArray(kidsObj);
            if (kids is null) return;
            foreach (var kid in kids.Items) {
                var d = ResolveDict(kid);
                if (d is null) { continue; }
                TraversePagesNode(d, outList);
            }
        }
    }

    private bool IsLikelyPage(PdfDictionary d) {
        // Heuristic when /Type is missing: leaf node has Contents, and page data can come from itself or inherited /Pages nodes.
        bool hasContents = d.Items.ContainsKey("Contents");
        bool hasRes = d.Items.ContainsKey("Resources") || HasInheritedValue(d, "Resources");
        bool hasMedia = HasMedia(d) || HasInheritedValue(d, "MediaBox") || HasInheritedValue(d, "CropBox");
        bool hasKids = ResolveArray(d.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null) is not null;
        return !hasKids && hasContents && (hasRes || hasMedia);
    }

    private void TraversePagesNodeDeepLimited(PdfDictionary node, HashSet<PdfDictionary> visitedNodes, HashSet<int> visitedPages, List<PdfReadPage> outList, int? limit) {
        if (!visitedNodes.Add(node)) {
            return;
        }

        var type = node.Get<PdfName>("Type")?.Name;
        if (type == "Page" || (type is null && IsLikelyPage(node))) {
            int objNum = FindObjectNumberFor(node);
            if (objNum > 0 && visitedPages.Add(objNum)) {
                if (type == "Page" || HasMedia(node) || HasInheritedValue(node, "MediaBox") || HasInheritedValue(node, "CropBox")) {
                    outList.Add(new PdfReadPage(objNum, node, _objects));
                }
            }
            return;
        }
        var kids = ResolveArray(node.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null);
        if (kids is null) return;
        foreach (var kid in kids.Items) {
            if (limit.HasValue && outList.Count >= limit.Value) return;
            var d = ResolveDict(kid);
            if (d is null) { continue; }
            var t = d.Get<PdfName>("Type")?.Name;
            if (t == "Pages" || (t is null && ResolveArray(d.Items.TryGetValue("Kids", out var dKidsObj) ? dKidsObj : null) is not null)) TraversePagesNodeDeepLimited(d, visitedNodes, visitedPages, outList, limit);
            else if ((t == "Page" || IsLikelyPage(d) || IsLeafPageByParent(d)) &&
                     (t == "Page" || HasMedia(d) || HasInheritedValue(d, "MediaBox") || HasInheritedValue(d, "CropBox"))) {
                int on = FindObjectNumberFor(d);
                if (on > 0 && visitedPages.Add(on)) {
                    outList.Add(new PdfReadPage(on, d, _objects));
                    if (limit.HasValue && outList.Count >= limit.Value) return;
                }
            }
        }
    }

    private HashSet<int> CollectReachableLeafCandidates(PdfDictionary pagesRoot) {
        var set = new HashSet<int>();
        var stack = new Stack<PdfDictionary>();
        stack.Push(pagesRoot);
        int guard = 0;
        while (stack.Count > 0 && guard++ < 10000) {
            var cur = stack.Pop();
            var kids = ResolveArray(cur.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null);
            if (kids is null) continue;
            foreach (var k in kids.Items) {
                var d = ResolveDict(k);
                if (d is null) continue;
                var t = d.Get<PdfName>("Type")?.Name;
                if (t == "Pages" || (t is null && ResolveArray(d.Items.TryGetValue("Kids", out var dKidsObj) ? dKidsObj : null) is not null)) stack.Push(d);
                else if (IsLikelyPage(d) || IsLeafPageByParent(d)) {
                    int on = FindObjectNumberFor(d);
                    if (on > 0) set.Add(on);
                }
            }
        }
        return set;
    }
    private bool IsLeafPageByParent(PdfDictionary d) {
        if (!IsLikelyPage(d)) return false;
        // Follow Parent chain up until /Pages or no parent
        PdfDictionary? current = d;
        int guard = 0;
        while (current is not null && guard++ < 100) {
            if (!current.Items.TryGetValue("Parent", out var p)) break;
            var parent = ResolveDict(p);
            if (parent is null) break;
            var type = parent.Get<PdfName>("Type")?.Name;
            if (type == "Pages") return true;
            current = parent;
        }
        return false;
    }

    private bool HasInheritedValue(PdfDictionary start, string key) {
        PdfDictionary? current = start;
        int guard = 0;
        while (current is not null && guard++ < 100) {
            if (current.Items.ContainsKey(key)) {
                return true;
            }

            if (!current.Items.TryGetValue("Parent", out var parentObj)) {
                break;
            }

            var parent = ResolveDict(parentObj);
            if (parent is null) {
                break;
            }

            current = parent;
        }

        return false;
    }

    private static bool HasMedia(PdfDictionary d) => d.Items.ContainsKey("MediaBox") || d.Items.ContainsKey("CropBox");

    private int FindExactObjectNumberFor(PdfDictionary dict) {
        foreach (var kv in _objects) if (ReferenceEquals(kv.Value.Value, dict)) return kv.Key;
        return 0;
    }

    private int FindObjectNumberFor(PdfDictionary dict) {
        foreach (var kv in _objects) if (ReferenceEquals(kv.Value.Value, dict)) return kv.Key;
        // As a fallback when dictionary was re-parsed separately, match by identity via a simple scan of Page objects
        foreach (var kv in _objects) if (kv.Value.Value is PdfDictionary d && d.Get<PdfName>("Type")?.Name == "Page") return kv.Key;
        return 0;
    }

    private string ToRaw() {
        // Reconstruct raw text for simple metadata extraction without reserialization; ok for small files.
        var sb = new StringBuilder();
        foreach (var kv in _objects.OrderBy(k => k.Key)) {
            sb.Append(kv.Key).Append(" 0 obj\n");
            if (kv.Value.Value is PdfStream s) {
                sb.Append("<< ");
                foreach (var d in s.Dictionary.Items) sb.Append('/').Append(d.Key).Append(' ').Append(' ').Append(' ');
                sb.Append(">>\nstream\n");
                sb.Append(PdfEncoding.Latin1GetString(s.Data)).Append("\nendstream\nendobj\n");
            } else {
                sb.Append("...\nendobj\n");
            }
        }
        sb.Append(_trailerRaw);
        return sb.ToString();
    }

    private PdfMetadata ExtractMetadata() {
        // Trailer has /Info N 0 R when present
        var m = System.Text.RegularExpressions.Regex.Match(_trailerRaw, @"/Info\s+(\d+)\s+0\s+R");
        if (!m.Success) return new PdfMetadata();
        if (!int.TryParse(m.Groups[1].Value, out int infoId)) return new PdfMetadata();
        if (!_objects.TryGetValue(infoId, out var infoObj) || infoObj.Value is not PdfDictionary dict) return new PdfMetadata();
        string? GetStr(string key) => dict.Get<PdfStringObj>(key)?.Value;
        return new PdfMetadata {
            Title = GetStr("Title"),
            Author = GetStr("Author"),
            Subject = GetStr("Subject"),
            Keywords = GetStr("Keywords")
        };
    }
}
