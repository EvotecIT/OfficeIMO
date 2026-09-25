namespace OfficeIMO.OpenDocument;

public sealed partial class OdtDocument {
    private NoteIndex? _noteIndex;

    internal void PrepareNoteIndexForMutation() => _ = GetNoteIndex();

    internal void RefreshNoteIndexAfterMutation() {
        NoteIndex? previous = _noteIndex;
        _noteIndex = null;
        if (previous == null || previous.ExternalXmlEditVersion != Package.ExternalXmlEditVersion) return;
        int footnotes = 0, endnotes = 0;
        bool changedContent = false, changedStyles = false;
        XDocument? stylesDocument = Package.ContainsEntry("styles.xml") ? GetXml("styles.xml") : null;
        foreach (XElement note in GetNotesInPackageOrder()) {
            OdtNoteKind? kind = NoteIndex.GetKind(note);
            if (!kind.HasValue) continue;
            int ordinal = kind == OdtNoteKind.Footnote ? ++footnotes : ++endnotes;
            if (!previous.IsGenerated(note)) continue;
            XElement? citation = note.Element(OdfNamespaces.Text + "note-citation");
            string value = ordinal.ToString(CultureInfo.InvariantCulture);
            if (citation == null || citation.Value == value) continue;
            citation.Value = value;
            if (note.Document == stylesDocument) changedStyles = true;
            else changedContent = true;
        }
        if (changedContent) MarkPartDirty("content.xml");
        if (changedStyles) MarkPartDirty("styles.xml");
    }

    internal OdtNote AddNote(XElement paragraph, string partPath, OdtNoteKind kind, string text) {
        ValidateNoteInsertion(kind, text);
        NoteIndex index = GetNoteIndex();

        string prefix = kind == OdtNoteKind.Footnote ? "ftng" : "endng";
        int ordinal = index.Count(kind) + 1;
        int idNumber = ordinal;
        string id;
        do { id = prefix + idNumber++.ToString(CultureInfo.InvariantCulture); }
        while (index.Ids.Contains(id));

        bool append = index.IsAfterLastNoteOfKind(paragraph, partPath, kind);
        OdtNote result = OdtNote.Create(this, kind, id, ordinal.ToString(CultureInfo.InvariantCulture), text, partPath);
        paragraph.Add(result.Element);
        if (append) {
            index.RecordAppend(result.Element, kind, id);
        } else {
            int next = 0;
            bool changedStyles = false;
            XDocument? stylesDocument = Package.ContainsEntry("styles.xml") ? GetXml("styles.xml") : null;
            foreach (XElement note in GetNotesInPackageOrder()) {
                if (NoteIndex.GetKind(note) != kind) continue;
                next++;
                if (note != result.Element && !index.IsGenerated(note)) continue;
                note.Element(OdfNamespaces.Text + "note-citation")!.Value = next.ToString(CultureInfo.InvariantCulture);
                if (note.Document == stylesDocument) changedStyles = true;
            }
            if (changedStyles) MarkPartDirty("styles.xml");
            _noteIndex = null;
        }
        MarkPartDirty(partPath);
        return result;
    }

    internal void ValidateNoteInsertion(OdtNoteKind kind, string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        NoteIndex index = GetNoteIndex();
        if (index.HasConfiguration(kind)) {
            throw new NotSupportedException("Adding notes to an ODT document with configured note numbering is not supported.");
        }
    }

    private NoteIndex GetNoteIndex() {
        if (_noteIndex == null || _noteIndex.ExternalXmlEditVersion != Package.ExternalXmlEditVersion)
            _noteIndex = new NoteIndex(this);
        return _noteIndex;
    }

    private IEnumerable<XElement> GetNotesInPackageOrder() => GetAllNotesInPackageOrder()
        .Where(note => !note.Ancestors(OdfNamespaces.Text + "tracked-changes").Any());

    private IEnumerable<XElement> GetAllNotesInPackageOrder() {
        foreach (XElement note in GetXml("content.xml").Descendants(OdfNamespaces.Text + "note")) yield return note;
        if (Package.ContainsEntry("styles.xml")) {
            foreach (XElement note in GetXml("styles.xml").Descendants(OdfNamespaces.Text + "note")) yield return note;
        }
    }

    private sealed class NoteIndex {
        private readonly OdtDocument _document;
        private readonly HashSet<string> _ids = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _generatedIds = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<OdtNoteKind> _configuredKinds = new HashSet<OdtNoteKind>();
        private readonly Dictionary<OdtNoteKind, int> _counts = new Dictionary<OdtNoteKind, int>();
        private readonly Dictionary<OdtNoteKind, XElement> _last = new Dictionary<OdtNoteKind, XElement>();

        internal NoteIndex(OdtDocument document) {
            _document = document;
            ExternalXmlEditVersion = document.Package.ExternalXmlEditVersion;
            foreach (string partPath in new[] { "content.xml", "styles.xml" }) {
                if (!document.Package.ContainsEntry(partPath)) continue;
                foreach (XElement config in document.GetXml(partPath).Descendants(OdfNamespaces.Text + "notes-configuration")) {
                    OdtNoteKind? kind = (string?)config.Attribute(OdfNamespaces.Text + "note-class") switch {
                        "footnote" => OdtNoteKind.Footnote,
                        "endnote" => OdtNoteKind.Endnote,
                        _ => null
                    };
                    if (kind.HasValue) _configuredKinds.Add(kind.Value);
                    else { _configuredKinds.Add(OdtNoteKind.Footnote); _configuredKinds.Add(OdtNoteKind.Endnote); }
                }
            }
            foreach (XElement note in document.GetAllNotesInPackageOrder()) {
                string? id = (string?)note.Attribute(OdfNamespaces.Text + "id");
                if (id != null) _ids.Add(id);
                if (id != null && note.Ancestors(OdfNamespaces.Text + "tracked-changes").Any() &&
                    IsGeneratedIdAndCitation(note, id)) _generatedIds.Add(id);
            }
            foreach (XElement note in document.GetNotesInPackageOrder()) {
                string? id = (string?)note.Attribute(OdfNamespaces.Text + "id");
                OdtNoteKind? kind = GetKind(note);
                if (!kind.HasValue) continue;
                int count = Count(kind.Value) + 1;
                _counts[kind.Value] = count;
                _last[kind.Value] = note;
                XElement? citation = note.Element(OdfNamespaces.Text + "note-citation");
                if (citation?.Attribute(OdfNamespaces.Text + "label") == null &&
                    citation?.Value == count.ToString(CultureInfo.InvariantCulture)) {
                    GeneratedCitations.Add(note);
                    if (id != null) _generatedIds.Add(id);
                }
            }
        }

        private static bool IsGeneratedIdAndCitation(XElement note, string id) {
            OdtNoteKind? kind = GetKind(note);
            string prefix = kind == OdtNoteKind.Footnote ? "ftn" : kind == OdtNoteKind.Endnote ? "endn" : string.Empty;
            if (prefix.Length == 0 || !id.StartsWith(prefix, StringComparison.Ordinal)) return false;
            string suffix = id.Substring(prefix.Length);
            bool markedGenerated = suffix.StartsWith("g", StringComparison.Ordinal);
            if (markedGenerated) suffix = suffix.Substring(1);
            XElement? citation = note.Element(OdfNamespaces.Text + "note-citation");
            return citation?.Attribute(OdfNamespaces.Text + "label") == null &&
                int.TryParse(suffix, NumberStyles.None, CultureInfo.InvariantCulture, out int number) &&
                number > 0 && (markedGenerated
                    ? int.TryParse(citation?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out int displayed) && displayed > 0
                    : citation?.Value == number.ToString(CultureInfo.InvariantCulture));
        }

        internal int ExternalXmlEditVersion { get; }
        internal HashSet<XElement> GeneratedCitations { get; } = new HashSet<XElement>();
        internal bool IsGenerated(XElement note) => GeneratedCitations.Contains(note) ||
            (string?)note.Attribute(OdfNamespaces.Text + "id") is string id && _generatedIds.Contains(id);
        internal HashSet<string> Ids => _ids;
        internal int Count(OdtNoteKind kind) => _counts.TryGetValue(kind, out int count) ? count : 0;
        internal bool HasConfiguration(OdtNoteKind kind) => _configuredKinds.Contains(kind);

        internal bool IsAfterLastNoteOfKind(XElement paragraph, string partPath, OdtNoteKind kind) {
            if (!_last.TryGetValue(kind, out XElement? last)) return true;
            string lastPart = last.Document == _document.GetXml("content.xml") ? "content.xml" : "styles.xml";
            if (lastPart != partPath) return partPath == "styles.xml";
            return last.AncestorsAndSelf().Contains(paragraph) ||
                XNode.DocumentOrderComparer.Compare(last, paragraph) < 0;
        }

        internal void RecordAppend(XElement note, OdtNoteKind kind, string id) {
            _ids.Add(id);
            _counts[kind] = Count(kind) + 1;
            _last[kind] = note;
            GeneratedCitations.Add(note);
            _generatedIds.Add(id);
        }

        internal static OdtNoteKind? GetKind(XElement note) =>
            (string?)note.Attribute(OdfNamespaces.Text + "note-class") switch {
                "footnote" => OdtNoteKind.Footnote,
                "endnote" => OdtNoteKind.Endnote,
                _ => null
            };
    }
}
