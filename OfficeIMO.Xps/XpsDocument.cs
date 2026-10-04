using OfficeIMO.Core.Internal;

namespace OfficeIMO.Xps;

/// <summary>A portable XPS/OpenXPS document sequence. Retains unmodified native package parts.</summary>
public sealed partial class XpsDocument {
    private readonly Dictionary<string, byte[]> _parts = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _types = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<XpsPage> _pages = new();
    private readonly Dictionary<string, Dictionary<string, int>> _linkTargets = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, int> _documentStarts = new(StringComparer.OrdinalIgnoreCase);
    private readonly XpsReadOptions _limits;
    private readonly string _sequence;
    private readonly string? _createdDocument;
    private XpsDocument(XpsFormat format, XpsReadOptions limits, string sequence, string? createdDocument = null) {
        Format = format; _limits = limits; _sequence = sequence; _createdDocument = createdDocument;
    }
    /// <summary>The markup dialect detected from the package relationship and sequence.</summary>
    public XpsFormat Format { get; }
    /// <summary>Pages in document-sequence order. Repeated references share one editable native page.</summary>
    public IReadOnlyList<XpsPage> Pages => _pages.AsReadOnly();
    /// <summary>Names of retained package parts, without a leading slash.</summary>
    public IReadOnlyCollection<string> PartNames => _parts.Keys.ToArray();
    /// <summary>Creates an empty native document with no external resources.</summary>
    public static XpsDocument Create(XpsFormat format = XpsFormat.OpenXps) {
        if (!Enum.IsDefined(typeof(XpsFormat), format)) throw new ArgumentOutOfRangeException(nameof(format));
        var result = new XpsDocument(format, new XpsReadOptions(), "FixedDocumentSequence.fdseq", "Documents/1/FixedDocument.fdoc");
        XNamespace ns = XpsPackage.Namespace(format);
        result.PutXml(result._sequence, new XElement(ns + "FixedDocumentSequence", new XElement(ns + "DocumentReference", new XAttribute("Source", "/" + result._createdDocument))), "fixeddocumentsequence");
        result.PutXml(result._createdDocument!, new XElement(ns + "FixedDocument"), "fixeddocument");
        result._parts.Add("_rels/.rels", XpsPackage.Serialize(new XElement(XpsPackage.Relationships + "Relationships",
            new XElement(XpsPackage.Relationships + "Relationship", new XAttribute("Id", "rStart"), new XAttribute("Type", XpsPackage.StartRelationship(format)), new XAttribute("Target", "/" + result._sequence)))));
        return result;
    }
    /// <summary>Loads from a file without retaining handles.</summary>
    public static XpsDocument Load(string path, XpsReadOptions? options = null, CancellationToken cancellationToken = default) {
        using var stream = File.OpenRead(path);
        return Load(stream, options, cancellationToken);
    }
    /// <summary>Loads from bytes with bounded package and XML expansion.</summary>
    public static XpsDocument Load(byte[] bytes, XpsReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, false);
        return Load(stream, options, cancellationToken);
    }
    /// <summary>Reads from the current position to EOF. The caller's stream remains open; non-seekable streams are supported.</summary>
    public static XpsDocument Load(Stream stream, XpsReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        var limits = (options ?? new XpsReadOptions()).Snapshot();
        using var input = new MemoryStream();
        var buffer = new byte[81920];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(buffer, 0, (int)Math.Min(buffer.Length, limits.MaximumInputBytes - input.Length + 1));
            if (read == 0) break;
            if (input.Length + read > limits.MaximumInputBytes) throw new InvalidDataException("XPS input byte limit exceeded.");
            input.Write(buffer, 0, read);
        }
        input.Position = 0;
        var scan = OfficeArchiveSafety.ScanZipCentralDirectory(input, input.Length, limits.MaximumParts, cancellationToken);
        if (!scan.IsValid || scan.LimitExceeded) throw new InvalidDataException("Invalid or excessive XPS ZIP directory.");
        var parts = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        using (var zip = new ZipArchive(input, ZipArchiveMode.Read, true)) {
            long expanded = 0;
            foreach (var entry in zip.Entries) {
                cancellationToken.ThrowIfCancellationRequested();
                if (entry.FullName.EndsWith("/", StringComparison.Ordinal)) continue;
                string name = XpsPackage.PartName(entry.FullName);
                if (entry.FullName.StartsWith("/", StringComparison.Ordinal) || parts.ContainsKey(name)) throw new InvalidDataException("Duplicate or absolute XPS ZIP part.");
                if (entry.Length > limits.MaximumPartBytes || entry.Length > limits.MaximumExpandedBytes - expanded) throw new InvalidDataException("XPS expanded byte limit exceeded.");
                expanded += entry.Length;
                using var part = entry.Open();
                parts.Add(name, OfficeArchiveSafety.ReadEntryBytes(part, entry.Length, limits.MaximumPartBytes));
            }
        }
        byte[] Required(string name) => parts.TryGetValue(name, out var data) ? data : throw new InvalidDataException("Missing XPS part: " + name);
        XElement rels = XpsPackage.Xml(Required("_rels/.rels"), limits, cancellationToken);
        if (rels.Name != XpsPackage.Relationships + "Relationships") throw new InvalidDataException("Invalid package relationships.");
        var starts = rels.Elements(XpsPackage.Relationships + "Relationship").Where(r => (string?)r.Attribute("Type") == XpsPackage.StartRelationship(XpsFormat.Xps) || (string?)r.Attribute("Type") == XpsPackage.StartRelationship(XpsFormat.OpenXps)).ToArray();
        if (starts.Length != 1 || ((string?)starts[0].Attribute("TargetMode") ?? "Internal") != "Internal") throw new InvalidDataException("An XPS package requires one internal fixed-representation relationship.");
        XpsFormat format = (string?)starts[0].Attribute("Type") == XpsPackage.StartRelationship(XpsFormat.Xps) ? XpsFormat.Xps : XpsFormat.OpenXps;
        string sequence = XpsPackage.Resolve("", (string?)starts[0].Attribute("Target") ?? "");
        var result = new XpsDocument(format, limits, sequence);
        foreach (var part in parts) result._parts.Add(part.Key, part.Value);
        result.ReadTypes(cancellationToken);
        XNamespace ns = XpsPackage.Namespace(format);
        XElement sequenceXml = result.RequiredXml(sequence, "FixedDocumentSequence", "fixeddocumentsequence", cancellationToken);
        if (sequenceXml.Elements().Any(e => e.Name != ns + "DocumentReference")) throw new NotSupportedException("Unsupported content in XPS document sequence.");
        var documents = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
        var pages = new Dictionary<string, XpsPage>(StringComparer.OrdinalIgnoreCase);
        var sequenceTargets = new Dictionary<string, int>(StringComparer.Ordinal);
        result._linkTargets.Add(sequence, sequenceTargets);
        foreach (var reference in sequenceXml.Elements(ns + "DocumentReference")) {
            cancellationToken.ThrowIfCancellationRequested();
            string docPart = XpsPackage.Resolve(sequence, (string?)reference.Attribute("Source") ?? "");
            bool firstDocumentReference = !documents.TryGetValue(docPart, out var doc);
            if (firstDocumentReference) {
                doc = result.RequiredXml(docPart, "FixedDocument", "fixeddocument", cancellationToken);
                documents.Add(docPart, doc);
                result._documentStarts.Add(docPart, doc.Elements(ns + "PageContent").Any() ? result._pages.Count : -1);
                result._linkTargets.Add(docPart, new Dictionary<string, int>(StringComparer.Ordinal));
            }
            if (doc!.Elements().Any(e => e.Name != ns + "PageContent")) throw new NotSupportedException("Unsupported content in XPS fixed document.");
            foreach (var pageRef in doc.Elements(ns + "PageContent")) {
                cancellationToken.ThrowIfCancellationRequested();
                if (result._pages.Count >= limits.MaximumPages) throw new InvalidDataException("XPS page limit exceeded.");
                string pagePart = XpsPackage.Resolve(docPart, (string?)pageRef.Attribute("Source") ?? "");
                // One native part has one editable backing, even when referenced repeatedly.
                if (!pages.TryGetValue(pagePart, out var page)) {
                    page = new XpsPage(result, pagePart, result.RequiredXml(pagePart, "FixedPage", "fixedpage", cancellationToken));
                    pages.Add(pagePart, page);
                }
                result._pages.Add(page);
                if (!firstDocumentReference) continue;
                foreach (var target in pageRef.Elements(ns + "PageContent.LinkTargets").Elements(ns + "LinkTarget")) {
                    string anchor = (string?)target.Attribute("Name") ?? throw new InvalidDataException("Missing link target name.");
                    var documentTargets = result._linkTargets[docPart];
                    if (!documentTargets.ContainsKey(anchor)) documentTargets.Add(anchor, result._pages.Count - 1);
                    if (!sequenceTargets.ContainsKey(anchor)) sequenceTargets.Add(anchor, result._pages.Count - 1);
                }
            }
        }
        return result;
    }
    internal int LinkTargetPage(string sourcePart, string? anchor) {
        if (string.Equals(sourcePart, _sequence, StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(anchor, NumberStyles.None, CultureInfo.InvariantCulture, out int number) && number > 0 && number <= _pages.Count) return number - 1;
        if (anchor != null && _linkTargets.TryGetValue(sourcePart, out var targets) && targets.TryGetValue(anchor, out int index)) return index;
        if (_documentStarts.TryGetValue(sourcePart, out int first)) return first < _pages.Count ? first : -1;
        if (string.Equals(sourcePart, _sequence, StringComparison.OrdinalIgnoreCase)) return _pages.Count > 0 ? 0 : -1;
        return _pages.FindIndex(p => string.Equals(p.PartName, sourcePart, StringComparison.OrdinalIgnoreCase));
    }
    internal byte[] Part(string name) => _parts.TryGetValue(name, out var bytes) ? bytes : throw new InvalidDataException("Missing XPS part: " + name);
    internal XElement ReadXml(string name, CancellationToken token) => XpsPackage.Xml(Part(name), _limits, token);
    internal string ContentType(string name) => _types.TryGetValue(name, out var type) ? type : throw new InvalidDataException("Missing content type: " + name);
    private XElement RequiredXml(string name, string root, string type, CancellationToken token) {
        if (ContentType(name) != XpsPackage.Type(type)) throw new InvalidDataException("Incorrect XPS content type: " + name);
        XElement xml = ReadXml(name, token);
        if (xml.Name != XName.Get(root, XpsPackage.Namespace(Format))) throw new InvalidDataException("Incorrect XPS root or mixed dialect: " + name);
        return xml;
    }
    private void ReadTypes(CancellationToken token) {
        XElement types = ReadXml("[Content_Types].xml", token);
        if (types.Name != XpsPackage.ContentTypes + "Types") throw new InvalidDataException("Invalid OPC content-types root.");
        var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in types.Elements(XpsPackage.ContentTypes + "Default")) defaults.Add((string?)item.Attribute("Extension") ?? "", (string?)item.Attribute("ContentType") ?? "");
        foreach (var name in _parts.Keys) if (defaults.TryGetValue(Path.GetExtension(name).TrimStart('.'), out var type)) _types[name] = type;
        var overrides = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in types.Elements(XpsPackage.ContentTypes + "Override")) {
            string name = XpsPackage.PartName((string?)item.Attribute("PartName") ?? "");
            if (!overrides.Add(name)) throw new InvalidDataException("Duplicate content type override.");
            _types[name] = (string?)item.Attribute("ContentType") ?? "";
        }
        foreach (string name in _parts.Keys.Where(n => n != "[Content_Types].xml")) _ = ContentType(name);
    }
    private void PutXml(string name, XElement xml, string type) { _parts[name] = XpsPackage.Serialize(xml); _types[name] = XpsPackage.Type(type); }
}
