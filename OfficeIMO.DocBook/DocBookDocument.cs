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

namespace OfficeIMO.DocBook;

/// <summary>
/// Source-preserving DocBook 4/5 document with a bounded typed common-structure API and extension access.
/// </summary>
public sealed partial class DocBookDocument {
    private readonly XDocument _xml;
    private readonly byte[]? _originalBytes;
    private readonly string? _originalText;
    private readonly string _originalXmlFingerprint;
    private bool _modified;

    private DocBookDocument(XDocument xml, DocBookProfile profile, byte[]? originalBytes, string? originalText) {
        _xml = xml; Profile = profile; _originalBytes = originalBytes; _originalText = originalText;
        _originalXmlFingerprint = GetXmlFingerprint(xml);
    }

    /// <summary>Exact writer and bounded-validation profile selected for this document.</summary>
    public DocBookProfile Profile { get; }
    /// <summary>Exact official schema identifiers associated with <see cref="Profile"/>.</summary>
    public DocBookSchemaProfile SchemaProfile => DocBookSchemaProfiles.Get(Profile);
    /// <summary>Document root kind.</summary>
    public DocBookDocumentKind Kind => RootElement.Name.LocalName == "book" ? DocBookDocumentKind.Book : DocBookDocumentKind.Article;
    /// <summary>True after a mutation through this API.</summary>
    public bool IsModified => HasChanges;
    /// <summary>Root element as a typed common node.</summary>
    public DocBookNode Root => new DocBookNode(this, RootElement);
    /// <summary>Underlying XML for extension inspection and advanced lossless editing.</summary>
    public XDocument Xml => _xml;
    internal XNamespace Namespace => RootElement.Name.Namespace;
    private XElement RootElement => _xml.Root ?? throw new InvalidDataException("The DocBook document has no root element.");

    /// <summary>Document title from the root metadata/title.</summary>
    public string? Title {
        get {
            XElement? info = FindInfo();
            return (info?.Element(Namespace + "title") ?? RootElement.Element(Namespace + "title"))?.Value;
        }
        set {
            XElement? info = FindInfo();
            XElement? title = info?.Element(Namespace + "title") ?? RootElement.Element(Namespace + "title");
            if (value == null) title?.Remove();
            else if (title == null) EnsureInfo().AddFirst(new XElement(Namespace + "title", value));
            else title.Value = value;
            MarkModified();
        }
    }

    /// <summary>Creates an article using the exact selected writer profile.</summary>
    public static DocBookDocument CreateArticle(DocBookProfile profile = DocBookProfile.DocBook52) => Create(DocBookDocumentKind.Article, profile);
    /// <summary>Creates a book using the exact selected writer profile.</summary>
    public static DocBookDocument CreateBook(DocBookProfile profile = DocBookProfile.DocBook52) => Create(DocBookDocumentKind.Book, profile);

    private static DocBookDocument Create(DocBookDocumentKind kind, DocBookProfile profile) {
        DocBookSchemaProfile schema = DocBookSchemaProfiles.Get(profile);
        string rootName = kind == DocBookDocumentKind.Article ? "article" : "book";
        XNamespace ns = schema.NamespaceUri;
        var root = new XElement(ns + rootName);
        if (profile == DocBookProfile.DocBook52) root.SetAttributeValue("version", "5.2");
        XDocumentType? type = profile == DocBookProfile.DocBook45
            ? new XDocumentType(rootName, schema.DtdPublicId, schema.DtdSystemId, null) : null;
        var xml = type == null
            ? new XDocument(new XDeclaration("1.0", "utf-8", null), root)
            : new XDocument(new XDeclaration("1.0", "utf-8", null), type, root);
        return new DocBookDocument(xml, profile, null, null) { _modified = true };
    }

    /// <summary>Parses DocBook text using secure, bounded XML settings.</summary>
    public static DocBookDocument Parse(string text, DocBookReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        options ??= new DocBookReadOptions(); options.Validate();
        if (text.Length > options.MaxCharacters) throw new InvalidDataException("DocBook input exceeds MaxCharacters.");
        cancellationToken.ThrowIfCancellationRequested();
        using var source = new StringReader(text);
        using var reader = CreateLimitingReader(XmlReader.Create(source, CreateSettings(options)), options, cancellationToken);
        XDocument xml = XDocument.Load(reader, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
        cancellationToken.ThrowIfCancellationRequested();
        DocBookProfile profile = DetectAndValidateShape(xml, options, cancellationToken);
        return new DocBookDocument(xml, profile, null, text);
    }

    /// <summary>Loads a DocBook file.</summary>
    public static DocBookDocument Load(string path, DocBookReadOptions? options = null, CancellationToken cancellationToken = default) {
        using var stream = File.OpenRead(path ?? throw new ArgumentNullException(nameof(path)));
        return Load(stream, options, cancellationToken);
    }

    /// <summary>Loads DocBook from a caller-owned stream while preserving seekable position.</summary>
    public static DocBookDocument Load(Stream stream, DocBookReadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new DocBookReadOptions(); options.Validate();
        return ParseBytes(OfficeStreamReader.ReadAllBytes(stream, cancellationToken, options.MaxInputBytes), options, cancellationToken);
    }

    /// <summary>Loads DocBook asynchronously from a caller-owned stream.</summary>
    public static async Task<DocBookDocument> LoadAsync(Stream stream, DocBookReadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new DocBookReadOptions(); options.Validate();
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, options.MaxInputBytes).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return ParseBytes(bytes, options, cancellationToken);
    }

    /// <summary>Adds a top-level section, creating a chapter container first for a book.</summary>
    public DocBookNode AddSection(string title) => Root.AddSection(title);
    /// <summary>Adds a top-level paragraph, creating a chapter container first for a book.</summary>
    public DocBookNode AddParagraph(string text) => Root.AddParagraph(text);

    /// <summary>
    /// Validates the bounded OfficeIMO common-structure profile. The returned result explicitly does not represent
    /// a complete external DTD, RELAX NG, Schematron, XInclude, assembly, or vocabulary-extension validation run.
    /// </summary>
    public DocBookValidationResult Validate(DocBookValidationOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new DocBookValidationOptions();
        options.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        var diagnostics = new DocBookDiagnosticCollector(options.MaxDetailedDiagnosticsPerCode);
        XElement? root = _xml.Root;
        if (root == null) {
            diagnostics.Add(new DocBookDiagnostic("DB001", DocBookDiagnosticSeverity.Error,
                "The document requires an article or book root element.", "/"));
            return new DocBookValidationResult(SchemaProfile, diagnostics.ToArray());
        }
        string rootName = root.Name.LocalName;
        if (rootName != "article" && rootName != "book") {
            diagnostics.Add(new DocBookDiagnostic("DB001", DocBookDiagnosticSeverity.Error, "The root must be article or book.", "/" + rootName));
        }
        if (Profile == DocBookProfile.DocBook52) {
            if (root.Name.NamespaceName != DocBookSchemaProfiles.DocBook52.NamespaceUri) {
                diagnostics.Add(new DocBookDiagnostic("DB002", DocBookDiagnosticSeverity.Error, "DocBook 5 requires the DocBook namespace.", "/" + rootName));
            }
            string? version = (string?)root.Attribute("version");
            if (version != "5.2") diagnostics.Add(new DocBookDiagnostic("DB003", DocBookDiagnosticSeverity.Warning,
                $"The document declares DocBook '{version ?? "unspecified"}'; OfficeIMO writes and validates the exact 5.2 common-structure profile.", "/" + rootName + "/@version"));
        } else if (root.Name.NamespaceName.Length != 0) {
            diagnostics.Add(new DocBookDiagnostic("DB004", DocBookDiagnosticSeverity.Error, "DocBook 4.5 elements must be unqualified.", "/" + rootName));
        } else if (_xml.DocumentType == null || _xml.DocumentType.Name != rootName ||
                   _xml.DocumentType.PublicId != DocBookSchemaProfiles.DocBook45.DtdPublicId ||
                   _xml.DocumentType.SystemId != DocBookSchemaProfiles.DocBook45.DtdSystemId) {
            diagnostics.Add(new DocBookDiagnostic("DB005", DocBookDiagnosticSeverity.Warning,
                "The document is read using the 4.5 common-structure profile but does not declare the exact DocBook XML 4.5 DTD identifiers.", "/" + rootName));
        }

        int position = 0;
        foreach (XElement element in root.DescendantsAndSelf()) {
            cancellationToken.ThrowIfCancellationRequested();
            string path = "/" + rootName + "//* [" + (++position) + "]";
            DocBookNodeKind kind = DocBookNames.GetKind(element.Name, Namespace);
            string localName = element.Name.LocalName;
            if (element != root && kind == DocBookNodeKind.Unknown) {
                diagnostics.Add(new DocBookDiagnostic("DB010", DocBookDiagnosticSeverity.Info,
                    $"Extension element '{element.Name}' is preserved but is outside the typed common-structure API.", path));
            }
            if (element != root && element.Name.Namespace == Namespace &&
                 (Profile == DocBookProfile.DocBook52 &&
                 (localName == "ulink" || kind == DocBookNodeKind.Info && localName != "info" ||
                  localName == "sect1" || localName == "sect2" || localName == "sect3" || localName == "sect4" || localName == "sect5") ||
                 Profile == DocBookProfile.DocBook45 && localName == "info")) {
                diagnostics.Add(new DocBookDiagnostic("DB014", DocBookDiagnosticSeverity.Error,
                    $"{localName} is not a typed element in the selected {Profile} common-structure profile.", path));
            }
            XElement? parent = element.Parent;
            DocBookNodeKind parentKind = parent == null ? DocBookNodeKind.Unknown : DocBookNames.GetKind(parent.Name, Namespace);
            string? expectedInfoName = parent == null ? null : GetComponentInfoElementName(parent);
            bool invalidInfoParent = kind == DocBookNodeKind.Info && element != root &&
                (expectedInfoName != null
                    ? !string.Equals(localName, expectedInfoName, StringComparison.Ordinal)
                    : parentKind != DocBookNodeKind.Unknown);
            bool invalidParent = kind == DocBookNodeKind.TableGroup && parentKind != DocBookNodeKind.Table ||
                (kind == DocBookNodeKind.TableHead || kind == DocBookNodeKind.TableBody) && parentKind != DocBookNodeKind.TableGroup ||
                kind == DocBookNodeKind.Row && parentKind != DocBookNodeKind.TableHead && parentKind != DocBookNodeKind.TableBody &&
                    parent?.Name != Namespace + "tfoot" ||
                kind == DocBookNodeKind.Entry && parentKind != DocBookNodeKind.Row ||
                kind == DocBookNodeKind.Author && parentKind != DocBookNodeKind.Info && parent?.Name != Namespace + "authorgroup" ||
                kind == DocBookNodeKind.ListItem && parentKind != DocBookNodeKind.ItemizedList &&
                    parentKind != DocBookNodeKind.OrderedList &&
                    !(parent?.Name == Namespace + "varlistentry" &&
                      parent?.Parent is XElement variableList &&
                      DocBookNames.GetKind(variableList.Name, Namespace) == DocBookNodeKind.VariableList) ||
                invalidInfoParent ||
                Kind == DocBookDocumentKind.Book && ReferenceEquals(parent, root) && element.Name.Namespace == Namespace &&
                    !IsAllowedBookRootChild(localName);
            if (invalidParent) {
                diagnostics.Add(new DocBookDiagnostic("DB015", DocBookDiagnosticSeverity.Error,
                    $"{localName} is not under a supported common-structure parent.", path));
            }
            XName xlinkHref = XName.Get("href", "http://www.w3.org/1999/xlink");
            if (kind == DocBookNodeKind.Link &&
                (element.Attribute("href") != null ||
                 Profile == DocBookProfile.DocBook52 && element.Attribute("url") != null ||
                 Profile == DocBookProfile.DocBook45 && localName == "ulink" && string.IsNullOrWhiteSpace((string?)element.Attribute("url")) ||
                 Profile == DocBookProfile.DocBook45 && localName != "ulink" && element.Attribute("url") != null ||
                 Profile == DocBookProfile.DocBook45 && element.Attribute(xlinkHref) != null)) {
                diagnostics.Add(new DocBookDiagnostic("DB016", DocBookDiagnosticSeverity.Error,
                    $"{localName} uses a link target attribute outside the selected {Profile} common-structure profile.", path));
            }
            if (kind == DocBookNodeKind.CrossReference) {
                if (string.IsNullOrWhiteSpace((string?)element.Attribute("linkend"))) {
                    diagnostics.Add(new DocBookDiagnostic("DB017", DocBookDiagnosticSeverity.Error,
                        "xref requires a nonblank linkend target in the bounded common-structure profile.", path));
                }
                if (element.Nodes().Any()) {
                    diagnostics.Add(new DocBookDiagnostic("DB017", DocBookDiagnosticSeverity.Error,
                        "xref must be empty in the bounded common-structure profile.", path));
                }
                if (element.Attribute("href") != null || element.Attribute("url") != null || element.Attribute(xlinkHref) != null) {
                    diagnostics.Add(new DocBookDiagnostic("DB016", DocBookDiagnosticSeverity.Error,
                        $"{localName} uses a link target attribute outside the selected {Profile} common-structure profile.", path));
                }
            }
            if (kind == DocBookNodeKind.Author && element.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value))) {
                diagnostics.Add(new DocBookDiagnostic("DB018", DocBookDiagnosticSeverity.Error,
                    "author text must be contained by a personname element in the bounded common-structure profile.", path));
            }
            if ((kind == DocBookNodeKind.Section || kind == DocBookNodeKind.Figure ||
                 kind == DocBookNodeKind.Table && element.Name.LocalName == "table") &&
                !element.Elements(Namespace + "title").Any()) {
                diagnostics.Add(new DocBookDiagnostic("DB011", DocBookDiagnosticSeverity.Warning,
                    $"{element.Name.LocalName} has no title in the bounded common-structure profile.", path));
            }
            if (kind == DocBookNodeKind.ListItem && !element.Elements().Any()) {
                diagnostics.Add(new DocBookDiagnostic("DB012", DocBookDiagnosticSeverity.Error, "listitem must contain content.", path));
            }
            if ((kind == DocBookNodeKind.TableHead || kind == DocBookNodeKind.TableBody || localName == "tfoot") &&
                !element.Elements(Namespace + "row").Any()) {
                diagnostics.Add(new DocBookDiagnostic("DB012", DocBookDiagnosticSeverity.Error,
                    $"{localName} must contain at least one row.", path));
            }
            if (kind == DocBookNodeKind.Row && !element.Elements(Namespace + "entry").Any()) {
                diagnostics.Add(new DocBookDiagnostic("DB012", DocBookDiagnosticSeverity.Error,
                    "row must contain at least one entry.", path));
            }
            if (kind == DocBookNodeKind.ImageData && string.IsNullOrWhiteSpace((string?)element.Attribute("fileref"))) {
                diagnostics.Add(new DocBookDiagnostic("DB013", DocBookDiagnosticSeverity.Error, "imagedata requires fileref.", path));
            }
        }
        return new DocBookValidationResult(SchemaProfile, diagnostics.ToArray());
    }

    /// <summary>Returns DocBook XML, preserving the exact unchanged source by default.</summary>
    public string ToDocBook(DocBookWriteOptions? options = null) {
        options ??= new DocBookWriteOptions();
        if (!HasChanges && options.PreserveUnchangedSource) {
            if (_originalText != null) return _originalText;
            if (_originalBytes != null) return OfficeXmlTextEncoding.Decode(_originalBytes, _xml.Declaration?.Encoding);
        }
        return Encoding.UTF8.GetString(Serialize(options));
    }

    /// <summary>Writes to a caller-owned stream.</summary>
    public void Write(Stream destination, DocBookWriteOptions? options = null) => OfficeStreamWriter.WriteAllBytes(destination, GetBytes(options));
    /// <summary>Writes asynchronously to a caller-owned stream.</summary>
    public Task WriteAsync(Stream destination, DocBookWriteOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeStreamWriter.WriteAllBytesAsync(destination, GetBytes(options), cancellationToken);
    /// <summary>Saves using an atomic same-directory file commit.</summary>
    public void Save(string path, DocBookWriteOptions? options = null) => OfficeFileCommit.WriteAllBytes(path, GetBytes(options));
    /// <summary>Saves asynchronously using an atomic same-directory file commit.</summary>
    public Task SaveAsync(string path, DocBookWriteOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeFileCommit.WriteAllBytesAsync(path, GetBytes(options), cancellationToken: cancellationToken);

    internal void MarkModified() => _modified = true;

    private XElement? FindInfo() {
        if (Profile == DocBookProfile.DocBook52) return RootElement.Element(Namespace + "info");
        return RootElement.Element(Namespace + (Kind == DocBookDocumentKind.Article ? "articleinfo" : "bookinfo"));
    }

    private XElement EnsureInfo() {
        XElement? info = FindInfo();
        if (info != null) return info;
        string name = Profile == DocBookProfile.DocBook52 ? "info" : Kind == DocBookDocumentKind.Article ? "articleinfo" : "bookinfo";
        info = new XElement(Namespace + name); RootElement.AddFirst(info); return info;
    }

    internal XElement ResolveTypedContentParent(XElement requestedParent, string localName) {
        if (localName == "author") {
            if (ReferenceEquals(requestedParent, RootElement)) return EnsureInfo();
            string? infoName = GetComponentInfoElementName(requestedParent);
            if (infoName != null) {
                XElement? info = requestedParent.Element(Namespace + infoName);
                if (info == null) {
                    info = new XElement(Namespace + infoName);
                    requestedParent.AddFirst(info);
                    MarkModified();
                }
                return info;
            }
        }
        if (Kind != DocBookDocumentKind.Book || !ReferenceEquals(requestedParent, RootElement) || IsAllowedBookRootChild(localName)) {
            return requestedParent;
        }
        XElement? chapter = RootElement.Elements(Namespace + "chapter").FirstOrDefault();
        if (chapter != null) return chapter;
        chapter = new XElement(Namespace + "chapter",
            new XElement(Namespace + "title", "Content"));
        RootElement.Add(chapter);
        MarkModified();
        return chapter;
    }

    internal string? GetComponentInfoElementName(XElement component) {
        string localName = component.Name.LocalName;
        bool isRoot = ReferenceEquals(component, RootElement);
        bool isSection = DocBookNames.GetKind(component.Name, Namespace) == DocBookNodeKind.Section;
        bool isSupportedComponent = localName == "chapter" || localName == "appendix" || localName == "article" ||
            localName == "bibliography" || localName == "glossary" || localName == "index" || localName == "part" ||
            localName == "preface" || localName == "reference" || localName == "setindex";
        if (!isRoot && !isSection && !isSupportedComponent) return null;
        if (Profile == DocBookProfile.DocBook52) return "info";
        switch (localName) {
            case "article": return "articleinfo";
            case "book": return "bookinfo";
            case "section": return "sectioninfo";
            case "sect1": return "sect1info";
            case "sect2": return "sect2info";
            case "sect3": return "sect3info";
            case "sect4": return "sect4info";
            case "sect5": return "sect5info";
            default: return localName + "info";
        }
    }

    private static bool IsAllowedBookRootChild(string localName) {
        switch (localName) {
            case "info":
            case "bookinfo":
            case "title":
            case "subtitle":
            case "titleabbrev":
            case "dedication":
            case "toc":
            case "lot":
            case "glossary":
            case "bibliography":
            case "preface":
            case "chapter":
            case "reference":
            case "part":
            case "article":
            case "appendix":
            case "index":
            case "setindex":
            case "colophon":
                return true;
            default:
                return false;
        }
    }

    private byte[] GetBytes(DocBookWriteOptions? options) {
        options ??= new DocBookWriteOptions();
        if (!HasChanges && options.PreserveUnchangedSource && _originalBytes != null) return (byte[])_originalBytes.Clone();
        return Serialize(options);
    }

    private byte[] Serialize(DocBookWriteOptions options) {
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

    private static DocBookDocument ParseBytes(byte[] bytes, DocBookReadOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        using var memory = new MemoryStream(bytes, writable: false);
        using var reader = CreateLimitingReader(XmlReader.Create(memory, CreateSettings(options)), options, cancellationToken);
        XDocument xml = XDocument.Load(reader, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
        cancellationToken.ThrowIfCancellationRequested();
        DocBookProfile profile = DetectAndValidateShape(xml, options, cancellationToken);
        return new DocBookDocument(xml, profile, (byte[])bytes.Clone(), null);
    }

    private static XmlReaderSettings CreateSettings(DocBookReadOptions options) => new XmlReaderSettings {
        DtdProcessing = DtdProcessing.Parse,
        XmlResolver = null,
        MaxCharactersInDocument = options.MaxCharacters,
        MaxCharactersFromEntities = options.MaxCharactersFromEntities,
        IgnoreComments = false,
        IgnoreProcessingInstructions = false
    };

    private static XmlReader CreateLimitingReader(XmlReader reader, DocBookReadOptions options, CancellationToken cancellationToken) =>
        new OfficeXmlLimitingReader(reader, "DocBook", options.MaxDepth, options.MaxElements, options.MaxAttributes, cancellationToken);

    private static DocBookProfile DetectAndValidateShape(XDocument xml, DocBookReadOptions options, CancellationToken cancellationToken) {
        XElement? root = xml.Root;
        if (root == null || (root.Name.LocalName != "article" && root.Name.LocalName != "book"))
            throw new InvalidDataException("The DocBook root must be article or book.");
        DocBookProfile profile;
        if (root.Name.NamespaceName == DocBookSchemaProfiles.DocBook52.NamespaceUri) profile = DocBookProfile.DocBook52;
        else if (root.Name.NamespaceName.Length == 0) profile = DocBookProfile.DocBook45;
        else throw new InvalidDataException($"Unsupported DocBook namespace '{root.Name.NamespaceName}'.");
        RejectUnsupportedEntityDeclarations(xml.DocumentType?.InternalSubset);

        int elements = 0, attributes = 0;
        var stack = new Stack<Tuple<XElement, int>>(); stack.Push(Tuple.Create(root, 0));
        while (stack.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            Tuple<XElement, int> entry = stack.Pop();
            if (entry.Item2 > options.MaxDepth) throw new InvalidDataException("DocBook input exceeds MaxDepth.");
            if (++elements > options.MaxElements) throw new InvalidDataException("DocBook input exceeds MaxElements.");
            attributes = checked(attributes + entry.Item1.Attributes().Count());
            if (attributes > options.MaxAttributes) throw new InvalidDataException("DocBook input exceeds MaxAttributes.");
            foreach (XElement child in entry.Item1.Elements().Reverse()) stack.Push(Tuple.Create(child, entry.Item2 + 1));
        }
        return profile;
    }

    private static void RejectUnsupportedEntityDeclarations(string? internalSubset) {
        if (string.IsNullOrEmpty(internalSubset)) return;
        for (int index = 0; index < internalSubset!.Length;) {
            if (Matches(index, "<?")) {
                int processingInstructionEnd = internalSubset.IndexOf("?>", index + 2, StringComparison.Ordinal);
                index = processingInstructionEnd < 0 ? internalSubset.Length : processingInstructionEnd + 2;
                continue;
            }
            if (Matches(index, "<!--")) {
                int commentEnd = internalSubset.IndexOf("-->", index + 4, StringComparison.Ordinal);
                index = commentEnd < 0 ? internalSubset.Length : commentEnd + 3;
                continue;
            }
            char current = internalSubset[index];
            if (current == '\'' || current == '"') {
                int literalEnd = internalSubset.IndexOf(current, index + 1);
                index = literalEnd < 0 ? internalSubset.Length : literalEnd + 1;
                continue;
            }
            if (!Matches(index, "<!ENTITY")) {
                index++;
                continue;
            }
            int cursor = index + 8;
            SkipWhitespace(ref cursor);
            if (cursor < internalSubset.Length && internalSubset[cursor] == '%') {
                throw UnsupportedEntity();
            }
            while (cursor < internalSubset.Length && !IsXmlWhitespace(internalSubset[cursor])) cursor++;
            SkipWhitespace(ref cursor);
            if (Matches(cursor, "SYSTEM") || Matches(cursor, "PUBLIC")) throw UnsupportedEntity();
            index = cursor;
        }

        void SkipWhitespace(ref int cursor) {
            while (cursor < internalSubset.Length && IsXmlWhitespace(internalSubset[cursor])) cursor++;
        }
        bool Matches(int offset, string value) => offset >= 0 && offset <= internalSubset.Length - value.Length &&
            string.CompareOrdinal(internalSubset, offset, value, 0, value.Length) == 0;
        static bool IsXmlWhitespace(char value) => value == ' ' || value == '\t' || value == '\r' || value == '\n';
        static InvalidDataException UnsupportedEntity() => new InvalidDataException(
            "DocBook internal subsets may use bounded internal general entities, but external and parameter entity declarations are not supported.");
    }

    private bool HasChanges => _modified || !string.Equals(_originalXmlFingerprint, GetXmlFingerprint(_xml), StringComparison.Ordinal);

    private static string GetXmlFingerprint(XDocument xml) =>
        (xml.Declaration?.ToString() ?? string.Empty) + "\n" + xml.ToString(SaveOptions.DisableFormatting);
}
