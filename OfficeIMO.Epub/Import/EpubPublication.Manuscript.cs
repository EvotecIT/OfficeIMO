namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    internal sealed class ManuscriptEntries {
        private readonly EpubPublication _publication;
        private readonly List<ManuscriptEntry> _entries = new();
        private long _payloadBytes;

        internal ManuscriptEntries(EpubPublication publication) => _publication = publication;

        internal IReadOnlyList<ManuscriptEntry> Entries => _entries;

        internal void Add(ManuscriptEntry entry) {
            _publication.EnsureEntryBudget(_entries.Count + 1);
            if (entry.Payload.LongLength > GetRemainingEntryBytes())
                throw new InvalidDataException("Resource exceeds the publication's retained entry or expanded-byte limit.");
            _entries.Add(entry);
            _payloadBytes += entry.Payload.LongLength;
        }

        internal void AddChapter(string id, string path, XDocument document) {
            _publication.EnsureEntryBudget(_entries.Count + 1);
            long remaining = GetRemainingEntryBytes();
            if (remaining <= 0) throw new InvalidDataException("The manuscript exceeds the publication's retained expanded-byte limit.");
            ValidateContent(document, "application/xhtml+xml");
            Add(new ManuscriptEntry(id, path, "application/xhtml+xml", SerializeXml(document, remaining), inSpine: true));
        }

        private long GetRemainingEntryBytes() => Math.Min(_publication._maximumEntryBytes,
            _publication._maximumRetainedBytes - _publication._retainedBytes - _payloadBytes);
    }

    internal sealed class ManuscriptEntry {
        internal ManuscriptEntry(string id, string path, string mediaType, byte[] payload, bool inSpine = false) {
            Id = id;
            Path = path;
            MediaType = mediaType;
            Payload = payload;
            InSpine = inSpine;
        }

        internal string Id { get; }
        internal string Path { get; }
        internal string MediaType { get; }
        internal byte[] Payload { get; }
        internal bool InSpine { get; }
    }

    // The importer owns the entry bytes and the newly created publication until the batch succeeds.
    internal void AddManuscriptEntries(IReadOnlyList<ManuscriptEntry> entries, CancellationToken token) {
        EnsureEntryBudget(entries.Count);
        var ids = new HashSet<string>(Root.DescendantsAndSelf().Attributes("id").Select(attribute => attribute.Value), StringComparer.Ordinal);
        var paths = new HashSet<string>(_entries.Keys.Append(PackagePath).Select(path => path.Normalize(NormalizationForm.FormC)),
            StringComparer.OrdinalIgnoreCase);
        var items = new List<XElement>(entries.Count);
        var positions = new List<XElement>();
        long payloadDelta = 0;
        foreach (ManuscriptEntry entry in entries) {
            token.ThrowIfCancellationRequested();
            XmlConvert.VerifyNCName(entry.Id);
            if (!ids.Add(entry.Id)) throw new ArgumentException("Package id already exists: " + entry.Id);
            string path = VerifyContentPath(entry.Path);
            if (!paths.Add(path.Normalize(NormalizationForm.FormC))) throw new ArgumentException("Entry path already exists: " + path);
            RequireText(entry.MediaType, nameof(entry.MediaType));
            if (IsScriptMediaType(entry.MediaType)) throw new NotSupportedException("Script authoring is outside the EPUB writer contract.");
            if (entry.Payload.LongLength > _maximumEntryBytes) throw new InvalidDataException("Resource exceeds the publication's retained entry-byte limit.");
            payloadDelta += entry.Payload.LongLength;
            items.Add(new XElement(Opf + "item", new XAttribute("id", entry.Id), new XAttribute("href", RelativeHref(PackagePath, path)),
                new XAttribute("media-type", entry.MediaType)));
            if (entry.InSpine) positions.Add(new XElement(Opf + "itemref", new XAttribute("idref", entry.Id)));
        }
        EditPackageElement(Root, proposed => {
            proposed.Element(Opf + "manifest")!.Add(items.Select(item => new XElement(item)));
            proposed.Element(Opf + "spine")!.Add(positions.Select(position => new XElement(position)));
        }, payloadDelta);
        foreach (ManuscriptEntry entry in entries) _entries.Add(entry.Path, entry.Payload);
        _retainedBytes += payloadDelta;
        MarkChanged();
    }

    // Every target is an importer-generated chapter path. Build the EPUB 3 TOC using that index.
    internal void SetManuscriptNavigation(IEnumerable<EpubNavigationEntry> headings, IReadOnlyList<ManuscriptEntry> entries,
        CancellationToken token) {
        string path = NavigationPath();
        byte[] previous = _entries[path];
        long available = Math.Min(_maximumEntryBytes, _maximumRetainedBytes - (RetainedPayloadBytes - previous.LongLength));
        if (available <= 0) throw new InvalidDataException("Navigation exceeds the publication's retained-byte limits.");
        XDocument navigation = ParseXml(previous, _maximumEntryBytes);
        XElement list = navigation.Descendants(Html + "nav").Single().Element(Html + "ol")!;
        var chapters = new HashSet<string>(entries.Where(entry => entry.InSpine).Select(entry => entry.Path), StringComparer.Ordinal);
        list.Add(WriteEntries(headings, 1));
        byte[] payload = SerializeXml(navigation, available);
        long delta = payload.LongLength - previous.LongLength;
        EnsurePackageBudget(_package, delta);
        _entries[path] = payload;
        _retainedBytes += delta;
        MarkChanged();

        IEnumerable<XElement> WriteEntries(IEnumerable<EpubNavigationEntry> nodes, int depth) {
            if (depth > 64) throw new InvalidDataException("Navigation depth exceeds 64.");
            foreach (EpubNavigationEntry node in nodes) {
                token.ThrowIfCancellationRequested();
                RequireText(node.Label, nameof(node.Label));
                EpubReference target = EpubReference.Resolve("package.opf", "/" + node.Target);
                if (target.Kind != EpubReferenceKind.Container || target.ContainerPath == null || !chapters.Contains(target.ContainerPath))
                    throw new InvalidDataException("Manuscript navigation must reference a generated chapter: " + node.Target);
                string href = RelativeHref(path, target.ContainerPath);
                if (target.Fragment != null) href += "#" + Uri.EscapeDataString(target.Fragment);
                yield return new XElement(Html + "li", new XElement(Html + "a", new XAttribute("href", href), node.Label),
                    node.Children.Count == 0 ? null : new XElement(Html + "ol", WriteEntries(node.Children, depth + 1)));
            }
        }
    }
}
