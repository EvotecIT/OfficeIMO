namespace OfficeIMO.Xps;

public sealed partial class XpsDocument {
    /// <summary>Fixed documents in sequence order. Repeated references share the same editable document.</summary>
    public IReadOnlyList<XpsFixedDocument> Documents => _documents.AsReadOnly();

    /// <summary>Appends an empty fixed document to this sequence.</summary>
    public XpsFixedDocument AddDocument() => InsertDocument(_documents.Count);

    /// <summary>Inserts a new empty fixed document. Existing parts and relationships are retained.</summary>
    public XpsFixedDocument InsertDocument(int index) {
        CheckIndex(index, _documents.Count, true);
        string name = NewPartName("Documents/", "/FixedDocument.fdoc");
        var document = new XpsFixedDocument(this, name, new XElement(XName.Get("FixedDocument", XpsPackage.Namespace(Format))));
        InsertDocumentCore(index, document);
        return document;
    }
    /// <summary>Inserts another reference to a fixed document owned by this package, including a removed document.</summary>
    public void InsertDocument(int index, XpsFixedDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (document.Owner != this) throw new ArgumentException("The fixed document belongs to another package.", nameof(document));
        CheckIndex(index, _documents.Count, true);
        InsertDocumentCore(index, document);
    }
    private void InsertDocumentCore(int index, XpsFixedDocument document) {
        var sequence = new XElement(_sequenceMarkup);
        InsertChild(sequence, index, new XElement(sequence.Name.Namespace + "DocumentReference", new XAttribute("Source", "/" + document.PartName)));
        CommitStructure(sequence, new Dictionary<XpsFixedDocument, XElement> { [document] = document.Markup });
    }
    /// <summary>Removes one sequence reference, retaining its native parts, resources and relationships.</summary>
    public void RemoveDocumentAt(int index) {
        CheckIndex(index, _documents.Count, false);
        var sequence = new XElement(_sequenceMarkup); sequence.Elements().ElementAt(index).Remove();
        CommitStructure(sequence);
    }
    /// <summary>Moves one sequence reference to its final zero-based index, retaining reference attributes.</summary>
    public void MoveDocument(int sourceIndex, int destinationIndex) {
        CheckIndex(sourceIndex, _documents.Count, false); CheckIndex(destinationIndex, _documents.Count, false);
        var sequence = new XElement(_sequenceMarkup); var item = sequence.Elements().ElementAt(sourceIndex); item.Remove();
        InsertChild(sequence, destinationIndex, item); CommitStructure(sequence);
    }

    internal static void CheckIndex(int index, int count, bool insertion) {
        if (index < 0 || index > count || (!insertion && index == count)) throw new ArgumentOutOfRangeException(nameof(index));
    }
    internal static void InsertChild(XElement parent, int index, XElement item) {
        var next = parent.Elements().ElementAtOrDefault(index);
        if (next == null) parent.Add(item); else next.AddBeforeSelf(item);
    }
    internal string NewPartName(string prefix, string suffix) {
        for (int i = 1; i <= _limits.MaximumParts; i++) {
            string name = prefix + i.ToString(CultureInfo.InvariantCulture) + suffix;
            if (!_parts.ContainsKey(name)) return name;
        }
        throw new InvalidOperationException("No available XPS part name within the package limit.");
    }
    internal XpsPage ResolvePage(string source, XElement reference, XpsPage? pending = null) {
        string name = XpsPackage.Resolve(source, (string?)reference.Attribute("Source") ?? "");
        if (pending != null && string.Equals(name, pending.PartName, StringComparison.OrdinalIgnoreCase)) return pending;
        return _pageCache.TryGetValue(name, out var page) ? page : throw new InvalidDataException("Unknown page reference: " + name);
    }
    private void ReadStructure(CancellationToken token) {
        _sequenceMarkup = RequiredXml(_sequence, "FixedDocumentSequence", "fixeddocumentsequence", token);
        XNamespace ns = XpsPackage.Namespace(Format);
        if (_sequenceMarkup.Elements().Any(e => e.Name != ns + "DocumentReference")) throw new NotSupportedException("Unsupported content in XPS document sequence.");
        long pageReferences = 0;
        foreach (var reference in _sequenceMarkup.Elements()) {
            token.ThrowIfCancellationRequested();
            string name = XpsPackage.Resolve(_sequence, (string?)reference.Attribute("Source") ?? "");
            if (_documentCache.TryGetValue(name, out var existing)) {
                pageReferences += existing.Markup.Elements().Count();
                if (pageReferences > _limits.MaximumPages) throw new InvalidDataException("XPS page limit exceeded.");
                continue;
            }
            var xml = RequiredXml(name, "FixedDocument", "fixeddocument", token);
            if (xml.Elements().Any(e => e.Name != ns + "PageContent")) throw new NotSupportedException("Unsupported content in XPS fixed document.");
            pageReferences += xml.Elements().Count();
            if (pageReferences > _limits.MaximumPages) throw new InvalidDataException("XPS page limit exceeded.");
            _documentCache.Add(name, new XpsFixedDocument(this, name, xml));
            foreach (var pageRef in xml.Elements()) {
                token.ThrowIfCancellationRequested();
                string part = XpsPackage.Resolve(name, (string?)pageRef.Attribute("Source") ?? "");
                if (!_pageCache.ContainsKey(part)) _pageCache.Add(part, new XpsPage(this, part, RequiredXml(part, "FixedPage", "fixedpage", token)));
            }
        }
        ApplyIndex(BuildIndex(_sequenceMarkup, token: token));
    }
    private sealed class StructureIndex {
        internal readonly List<XpsFixedDocument> Documents = new();
        internal readonly List<XpsPage> Pages = new();
        internal readonly Dictionary<string, int> Starts = new(StringComparer.OrdinalIgnoreCase);
        internal readonly Dictionary<string, Dictionary<string, int>> Targets = new(StringComparer.OrdinalIgnoreCase);
    }
    private StructureIndex BuildIndex(XElement sequence, IReadOnlyDictionary<XpsFixedDocument, XElement>? changes = null, XpsPage? pending = null, CancellationToken token = default) {
        var index = new StructureIndex(); XNamespace ns = XpsPackage.Namespace(Format);
        var sequenceTargets = new Dictionary<string, int>(StringComparer.Ordinal); index.Targets.Add(_sequence, sequenceTargets);
        foreach (var reference in sequence.Elements()) {
            token.ThrowIfCancellationRequested();
            string part = XpsPackage.Resolve(_sequence, (string?)reference.Attribute("Source") ?? "");
            XpsFixedDocument document = _documentCache.TryGetValue(part, out var known) ? known
                : changes!.Keys.Single(d => string.Equals(d.PartName, part, StringComparison.OrdinalIgnoreCase));
            index.Documents.Add(document);
            XElement xml = changes != null && changes.TryGetValue(document, out var replacement) ? replacement : document.Markup;
            bool first = !index.Starts.ContainsKey(part);
            if (first) { index.Starts.Add(part, xml.HasElements ? index.Pages.Count : -1); index.Targets.Add(part, new Dictionary<string, int>(StringComparer.Ordinal)); }
            foreach (var pageRef in xml.Elements()) {
                token.ThrowIfCancellationRequested();
                if (index.Pages.Count >= _limits.MaximumPages) throw new InvalidDataException("XPS page limit exceeded.");
                index.Pages.Add(ResolvePage(part, pageRef, pending));
                if (!first) continue;
                foreach (var target in pageRef.Elements(ns + "PageContent.LinkTargets").Elements(ns + "LinkTarget")) {
                    string anchor = (string?)target.Attribute("Name") ?? throw new InvalidDataException("Missing link target name.");
                    if (!index.Targets[part].ContainsKey(anchor)) index.Targets[part].Add(anchor, index.Pages.Count - 1);
                    if (!sequenceTargets.ContainsKey(anchor)) sequenceTargets.Add(anchor, index.Pages.Count - 1);
                }
            }
        }
        return index;
    }
    private void ApplyIndex(StructureIndex index) {
        _documents.Clear(); _documents.AddRange(index.Documents);
        _pages.Clear(); _pages.AddRange(index.Pages);
        _documentStarts.Clear(); foreach (var item in index.Starts) _documentStarts.Add(item.Key, item.Value);
        _linkTargets.Clear(); foreach (var item in index.Targets) _linkTargets.Add(item.Key, item.Value);
    }
    // Prepare and validate every replacement before mutating native structure or public backing objects.
    internal void CommitStructure(XElement sequence, IReadOnlyDictionary<XpsFixedDocument, XElement>? changes = null, XpsPage? pending = null) {
        var index = BuildIndex(sequence, changes, pending);
        var navigation = PreserveNavigationTargets(index);
        var documentStructures = PreserveDocumentStructure(index);
        var storyBudget = new XpsStoryFragmentsReader.Budget(default);
        foreach (var page in index.Pages.Distinct()) {
            var structure = ReadPageContent(page, index.Pages.IndexOf(page), storyBudget);
            if (structure.HasUnresolvedNames)
                throw new InvalidDataException("Structural edits cannot retain unresolved native content references.");
        }
        var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) { [_sequence] = XpsPackage.Serialize(sequence) };
        foreach (var structure in documentStructures) replacements[structure.Key] = structure.Value;
        if (changes != null) foreach (var change in changes) replacements[change.Key.PartName] = XpsPackage.Serialize(change.Value);
        if (pending != null) replacements[pending.PartName] = pending.Serialize();
        foreach (var page in navigation) replacements[page.Key.PartName] = XpsPackage.Serialize(page.Value);
        var replacementTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (changes != null) foreach (var change in changes) replacementTypes[change.Key.PartName] = XpsPackage.Type("fixeddocument");
        if (pending != null) replacementTypes[pending.PartName] = XpsPackage.Type("fixedpage");
        foreach (var item in replacements) {
            if (item.Value.Length > _limits.MaximumPartBytes) throw new InvalidDataException("XPS part byte limit exceeded.");
            _ = XpsPackage.Xml(item.Value, _limits, default);
        }
        _ = PrepareOutput(default, replacements, replacementTypes, pending);
        foreach (var item in replacements) _parts[item.Key] = item.Value;
        _sequenceMarkup = sequence;
        if (changes != null) foreach (var change in changes) {
            change.Key.Markup = change.Value; _documentCache[change.Key.PartName] = change.Key; _types[change.Key.PartName] = XpsPackage.Type("fixeddocument");
        }
        if (pending != null) { _pageCache[pending.PartName] = pending; _types[pending.PartName] = XpsPackage.Type("fixedpage"); }
        foreach (var page in navigation) page.Key.ApplyMarkup(page.Value);
        ApplyIndex(index);
    }
    internal XElement ValidatePageMarkup(XElement markup) {
        byte[] bytes = XpsPackage.Serialize(markup);
        if (bytes.Length > _limits.MaximumPartBytes) throw new InvalidDataException("XPS part byte limit exceeded.");
        return XpsPackage.Xml(bytes, _limits, default);
    }
    internal void CommitDocument(XpsFixedDocument document, XElement markup, XpsPage? pending = null, bool attach = false) {
        var sequence = new XElement(_sequenceMarkup);
        if (attach) sequence.Add(new XElement(sequence.Name.Namespace + "DocumentReference", new XAttribute("Source", "/" + document.PartName)));
        CommitStructure(sequence, new Dictionary<XpsFixedDocument, XElement> { [document] = markup }, pending);
    }
    internal void CommitDocuments(Dictionary<XpsFixedDocument, XElement> changes) => CommitStructure(new XElement(_sequenceMarkup), changes);
}
