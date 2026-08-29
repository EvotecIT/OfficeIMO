using OfficeIMO.Core.Internal;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Opml;

/// <summary>A source-preserving OPML 1.0/2.0 document with typed editing helpers.</summary>
public sealed partial class OpmlDocument {
    private readonly XDocument _xml;
    private readonly byte[]? _originalBytes;
    private readonly string? _originalText;
    private readonly string _originalXmlFingerprint;
    private bool _modified;

    private OpmlDocument(XDocument xml, byte[]? originalBytes, string? originalText) {
        _xml = xml;
        _originalBytes = originalBytes;
        _originalText = originalText;
        _originalXmlFingerprint = GetXmlFingerprint(xml);
    }

    /// <summary>The declared OPML version text, including a preserved 1.1 declaration.</summary>
    public string DeclaredVersion {
        get => (string?)Root.Attribute("version") ?? string.Empty;
        set {
            Root.SetAttributeValue("version", value ?? throw new ArgumentNullException(nameof(value)));
            MarkModified();
        }
    }

    /// <summary>The effective supported OPML profile.</summary>
    public OpmlVersion Version => DeclaredVersion == "2.0" ? OpmlVersion.Opml20 : OpmlVersion.Opml10;
    /// <summary>True after a typed or extension mutation.</summary>
    public bool IsModified => HasChanges;
    /// <summary>Typed standard head values.</summary>
    public OpmlHead Head => new OpmlHead(this, HeadElement);
    /// <summary>Top-level outlines in document order.</summary>
    public IReadOnlyList<OpmlOutline> Outlines => BodyElement.Elements("outline").Select(e => new OpmlOutline(this, e)).ToArray();
    /// <summary>Underlying XML for advanced lossless extension inspection.</summary>
    public XDocument Xml => _xml;

    private XElement Root => _xml.Root ?? throw new InvalidDataException("The OPML document has no root element.");
    private XElement HeadElement => Root.Element("head") ?? throw new InvalidDataException("The OPML document has no head element.");
    private XElement BodyElement => Root.Element("body") ?? throw new InvalidDataException("The OPML document has no body element.");

    /// <summary>Creates an empty OPML document.</summary>
    public static OpmlDocument Create(OpmlVersion version = OpmlVersion.Opml20) {
        string declared;
        switch (version) {
            case OpmlVersion.Opml10: declared = "1.0"; break;
            case OpmlVersion.Opml20: declared = "2.0"; break;
            default: throw new ArgumentOutOfRangeException(nameof(version));
        }
        var xml = new XDocument(new XDeclaration("1.0", "utf-8", null),
            new XElement("opml", new XAttribute("version", declared), new XElement("head"), new XElement("body")));
        return new OpmlDocument(xml, null, null) { _modified = true };
    }

    /// <summary>Parses OPML text with secure, bounded XML settings.</summary>
    public static OpmlDocument Parse(string text, OpmlReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        options ??= new OpmlReadOptions();
        options.Validate();
        if (text.Length > options.MaxCharacters) throw new InvalidDataException("OPML input exceeds MaxCharacters.");
        cancellationToken.ThrowIfCancellationRequested();
        XDocument xml = ParseXml(new StringReader(text), options, cancellationToken);
        ValidateShapeAndLimits(xml, options, cancellationToken);
        return new OpmlDocument(xml, null, text);
    }

    /// <summary>Loads an OPML file.</summary>
    public static OpmlDocument Load(string path, OpmlReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        using var stream = File.OpenRead(path);
        return Load(stream, options, cancellationToken);
    }

    /// <summary>Loads OPML from a caller-owned stream without changing a seekable stream's position.</summary>
    public static OpmlDocument Load(Stream stream, OpmlReadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new OpmlReadOptions();
        options.Validate();
        byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, cancellationToken, options.MaxInputBytes);
        return ParseBytes(bytes, options, cancellationToken);
    }

    /// <summary>Loads OPML asynchronously from a caller-owned stream.</summary>
    public static async Task<OpmlDocument> LoadAsync(Stream stream, OpmlReadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new OpmlReadOptions();
        options.Validate();
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, options.MaxInputBytes).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return ParseBytes(bytes, options, cancellationToken);
    }

    /// <summary>Adds a top-level outline.</summary>
    public OpmlOutline AddOutline(string text) {
        var element = new XElement("outline", new XAttribute("text", text ?? throw new ArgumentNullException(nameof(text))));
        BodyElement.Add(element);
        MarkModified();
        return new OpmlOutline(this, element);
    }

    /// <summary>Validates the supported OPML profile without discarding extension content.</summary>
    public OpmlValidationResult Validate() => Validate(null, default);

    /// <summary>Validates the supported OPML profile while observing cancellation during semantic walks.</summary>
    public OpmlValidationResult Validate(CancellationToken cancellationToken) => Validate(null, cancellationToken);

    /// <summary>Validates the supported OPML profile with a bounded diagnostic budget.</summary>
    public OpmlValidationResult Validate(OpmlValidationOptions? options) => Validate(options, default);

    /// <summary>Validates the supported OPML profile with a bounded diagnostic budget and cancellation.</summary>
    public OpmlValidationResult Validate(OpmlValidationOptions? options, CancellationToken cancellationToken) {
        options ??= new OpmlValidationOptions();
        options.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        var diagnostics = new OpmlDiagnosticCollector(options.MaxDetailedDiagnosticsPerCode);
        XElement? root = _xml.Root;
        string declaredVersion = (string?)root?.Attribute("version") ?? string.Empty;
        OpmlVersion profile = declaredVersion == "2.0" ? OpmlVersion.Opml20 : OpmlVersion.Opml10;
        if (root == null || root.Name != "opml") {
            diagnostics.Add(new OpmlDiagnostic("OPML003", OpmlDiagnosticSeverity.Error,
                "The document root must be opml in no namespace.", "/"));
        }

        var rootChildrenList = new List<XElement>();
        if (root != null) {
            foreach (XElement child in root.Elements()) {
                cancellationToken.ThrowIfCancellationRequested();
                rootChildrenList.Add(child);
            }
        }
        XElement[] rootChildren = rootChildrenList.ToArray();
        XElement[] headElements = rootChildren.Where(element => element.Name == "head").ToArray();
        XElement[] bodyElements = rootChildren.Where(element => element.Name == "body").ToArray();
        if (headElements.Length != 1 || bodyElements.Length != 1) {
            diagnostics.Add(new OpmlDiagnostic("OPML004", OpmlDiagnosticSeverity.Error,
                "An OPML document requires exactly one head and one body element.", "/opml"));
        } else {
            if (Array.IndexOf(rootChildren, headElements[0]) > Array.IndexOf(rootChildren, bodyElements[0])) {
                diagnostics.Add(new OpmlDiagnostic("OPML005", OpmlDiagnosticSeverity.Error,
                    "The OPML head element must precede the body element.", "/opml"));
            }
        }

        if (declaredVersion != "1.0" && declaredVersion != "1.1" && declaredVersion != "2.0") {
            diagnostics.Add(new OpmlDiagnostic("OPML001", OpmlDiagnosticSeverity.Error,
                $"Unsupported OPML version '{declaredVersion}'.", "/opml/@version"));
        } else if (declaredVersion == "1.1") {
            diagnostics.Add(new OpmlDiagnostic("OPML002", OpmlDiagnosticSeverity.Info,
                "OPML 1.1 is interpreted using the OPML 1.0 profile.", "/opml/@version"));
        }

        int index = 0;
        IEnumerable<OpmlOutline> outlines = bodyElements.Length == 1
            ? bodyElements[0].Descendants("outline").Select(element => new OpmlOutline(this, element))
            : Enumerable.Empty<OpmlOutline>();
        foreach (OpmlOutline outline in outlines) {
            cancellationToken.ThrowIfCancellationRequested();
            string path = $"/opml/body//outline[{++index}]";
            if (outline.Element.Attribute("text") == null) {
                diagnostics.Add(new OpmlDiagnostic("OPML010", OpmlDiagnosticSeverity.Error,
                    "Every outline requires a text attribute.", path));
            }
            if (string.Equals(outline.Type, "rss", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(outline.XmlUrl)) {
                diagnostics.Add(new OpmlDiagnostic("OPML011", OpmlDiagnosticSeverity.Error,
                    "An rss subscription outline requires xmlUrl.", path));
            }
            if ((string.Equals(outline.Type, "link", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(outline.Type, "include", StringComparison.OrdinalIgnoreCase)) && string.IsNullOrWhiteSpace(outline.Url)) {
                diagnostics.Add(new OpmlDiagnostic("OPML012", OpmlDiagnosticSeverity.Error,
                    "A link/include outline requires url.", path));
            }
        }
        return new OpmlValidationResult(profile, diagnostics.ToArray());
    }

    /// <summary>Returns OPML text, preserving the exact input while unchanged by default.</summary>
    public string ToOpml(OpmlWriteOptions? options = null) {
        options ??= new OpmlWriteOptions();
        if (!HasChanges && options.PreserveUnchangedSource) {
            if (_originalText != null) return _originalText;
            if (_originalBytes != null) return OfficeXmlTextEncoding.Decode(_originalBytes, _xml.Declaration?.Encoding);
        }
        return Encoding.UTF8.GetString(Serialize(options));
    }

    /// <summary>Writes OPML to a caller-owned stream and rewinds a seekable destination.</summary>
    public void Write(Stream destination, OpmlWriteOptions? options = null) =>
        OfficeStreamWriter.WriteAllBytes(destination, GetBytes(options));

    /// <summary>Writes OPML asynchronously to a caller-owned stream.</summary>
    public Task WriteAsync(Stream destination, OpmlWriteOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeStreamWriter.WriteAllBytesAsync(destination, GetBytes(options), cancellationToken);

    /// <summary>Saves OPML through an atomic same-directory file commit.</summary>
    public void Save(string path, OpmlWriteOptions? options = null) => OfficeFileCommit.WriteAllBytes(path, GetBytes(options));

    /// <summary>Saves OPML asynchronously through an atomic same-directory file commit.</summary>
    public Task SaveAsync(string path, OpmlWriteOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeFileCommit.WriteAllBytesAsync(path, GetBytes(options), cancellationToken: cancellationToken);

    internal void MarkModified() => _modified = true;

    internal IEnumerable<OpmlOutline> Descendants() =>
        BodyElement.Descendants("outline").Select(e => new OpmlOutline(this, e));

    private byte[] GetBytes(OpmlWriteOptions? options) {
        options ??= new OpmlWriteOptions();
        if (!HasChanges && options.PreserveUnchangedSource && _originalBytes != null) return (byte[])_originalBytes.Clone();
        return Serialize(options);
    }

    private byte[] Serialize(OpmlWriteOptions options) {
        using var output = new MemoryStream();
        var settings = new XmlWriterSettings {
            Encoding = new UTF8Encoding(false),
            Indent = options.Indent,
            OmitXmlDeclaration = false,
            NewLineHandling = NewLineHandling.None
        };
        using (XmlWriter writer = XmlWriter.Create(output, settings)) _xml.Save(writer);
        return output.ToArray();
    }

    private static OpmlDocument ParseBytes(byte[] bytes, OpmlReadOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        using var memory = new MemoryStream(bytes, writable: false);
        using var reader = CreateLimitingReader(XmlReader.Create(memory, CreateSettings(options)), options, cancellationToken);
        XDocument xml = XDocument.Load(reader, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
        cancellationToken.ThrowIfCancellationRequested();
        ValidateShapeAndLimits(xml, options, cancellationToken);
        return new OpmlDocument(xml, (byte[])bytes.Clone(), null);
    }

    private static XDocument ParseXml(TextReader source, OpmlReadOptions options, CancellationToken cancellationToken) {
        using var reader = CreateLimitingReader(XmlReader.Create(source, CreateSettings(options)), options, cancellationToken);
        XDocument result = XDocument.Load(reader, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
        cancellationToken.ThrowIfCancellationRequested();
        return result;
    }

    private static XmlReaderSettings CreateSettings(OpmlReadOptions options) => new XmlReaderSettings {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        MaxCharactersInDocument = options.MaxCharacters,
        MaxCharactersFromEntities = 0,
        IgnoreComments = false,
        IgnoreProcessingInstructions = false
    };

    private static XmlReader CreateLimitingReader(XmlReader reader, OpmlReadOptions options, CancellationToken cancellationToken) =>
        new OfficeXmlLimitingReader(reader, "OPML", options.MaxDepth, options.MaxElements, options.MaxAttributes,
            cancellationToken, "outline", string.Empty, options.MaxOutlines, "MaxOutlines");

    private static void ValidateShapeAndLimits(XDocument xml, OpmlReadOptions options, CancellationToken cancellationToken) {
        XElement? root = xml.Root;
        if (root == null || root.Name != "opml") throw new InvalidDataException("The document root must be opml in no namespace.");
        if (root.Elements("head").Count() != 1 || root.Elements("body").Count() != 1) {
            throw new InvalidDataException("An OPML document requires exactly one head and one body element.");
        }
        XElement[] rootChildren = root.Elements().ToArray();
        if (Array.IndexOf(rootChildren, root.Element("head")!) > Array.IndexOf(rootChildren, root.Element("body")!)) {
            throw new InvalidDataException("The OPML head element must precede the body element.");
        }
        int elements = 0;
        int outlines = 0;
        int attributes = 0;
        var stack = new Stack<Tuple<XElement, int>>();
        stack.Push(Tuple.Create(root, 0));
        while (stack.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            Tuple<XElement, int> entry = stack.Pop();
            if (entry.Item2 > options.MaxDepth) throw new InvalidDataException("OPML input exceeds MaxDepth.");
            if (++elements > options.MaxElements) throw new InvalidDataException("OPML input exceeds MaxElements.");
            attributes = checked(attributes + entry.Item1.Attributes().Count());
            if (attributes > options.MaxAttributes) throw new InvalidDataException("OPML input exceeds MaxAttributes.");
            if (entry.Item1.Name == "outline" && ++outlines > options.MaxOutlines) throw new InvalidDataException("OPML input exceeds MaxOutlines.");
            foreach (XElement child in entry.Item1.Elements().Reverse()) stack.Push(Tuple.Create(child, entry.Item2 + 1));
        }
    }

    private bool HasChanges => _modified || !string.Equals(_originalXmlFingerprint, GetXmlFingerprint(_xml), StringComparison.Ordinal);

    private static string GetXmlFingerprint(XDocument xml) =>
        (xml.Declaration?.ToString() ?? string.Empty) + "\n" + xml.ToString(SaveOptions.DisableFormatting);
}
