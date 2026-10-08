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
    private XElement _sequenceMarkup = null!;
    private readonly List<XpsFixedDocument> _documents = new();
    private readonly Dictionary<string, XpsFixedDocument> _documentCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, XpsPage> _pageCache = new(StringComparer.OrdinalIgnoreCase);
    private XpsDocument(XpsFormat format, XpsReadOptions limits, string sequence) {
        Format = format; _limits = limits; _sequence = sequence;
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
        var result = new XpsDocument(format, new XpsReadOptions(), "FixedDocumentSequence.fdseq");
        XNamespace ns = XpsPackage.Namespace(format);
        result.PutXml(result._sequence, new XElement(ns + "FixedDocumentSequence", new XElement(ns + "DocumentReference", new XAttribute("Source", "/Documents/1/FixedDocument.fdoc"))), "fixeddocumentsequence");
        result.PutXml("Documents/1/FixedDocument.fdoc", new XElement(ns + "FixedDocument"), "fixeddocument");
        result._parts.Add("_rels/.rels", XpsPackage.Serialize(new XElement(XpsPackage.Relationships + "Relationships",
            new XElement(XpsPackage.Relationships + "Relationship", new XAttribute("Id", "rStart"), new XAttribute("Type", XpsPackage.StartRelationship(format)), new XAttribute("Target", "/" + result._sequence)))));
        result.ReadStructure(default);
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
        parts = OfficeOpcPieceAssembler.Assemble(parts, limits.MaximumPartBytes, cancellationToken);
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
        result.ReadStructure(cancellationToken);
        return result;
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
