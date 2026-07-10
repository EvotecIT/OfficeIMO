using OfficeIMO.Rtf.Syntax;
using OfficeIMO.Rtf.Writing;

namespace OfficeIMO.Rtf;

/// <summary>
/// Semantic Rich Text Format document model.
/// </summary>
public sealed partial class RtfDocument {
    private readonly List<RtfParagraph> _paragraphs = new List<RtfParagraph>();
    private readonly List<IRtfBlock> _blocks = new List<IRtfBlock>();
    private readonly List<RtfFont> _fonts = new List<RtfFont>();
    private readonly List<RtfColor> _colors = new List<RtfColor>();
    private readonly List<RtfStyle> _styles = new List<RtfStyle>();
    private readonly List<RtfListDefinition> _listDefinitions = new List<RtfListDefinition>();
    private readonly List<RtfListOverride> _listOverrides = new List<RtfListOverride>();
    private readonly List<RtfHeaderFooter> _headerFooters = new List<RtfHeaderFooter>();
    private readonly List<RtfNote> _notes = new List<RtfNote>();
    private readonly List<RtfSection> _sections = new List<RtfSection>();
    private readonly List<RtfUserProperty> _userProperties = new List<RtfUserProperty>();
    private readonly List<RtfDocumentVariable> _documentVariables = new List<RtfDocumentVariable>();
    private readonly List<RtfRevisionAuthor> _revisionAuthors = new List<RtfRevisionAuthor>();
    private readonly List<int> _revisionSaveIds = new List<int>();
    private readonly List<RtfFileReference> _fileReferences = new List<RtfFileReference>();
    private readonly List<RtfXmlNamespace> _xmlNamespaces = new List<RtfXmlNamespace>();

    private RtfDocument() {
    }

    /// <summary>Paragraphs contained by the document.</summary>
    public IReadOnlyList<RtfParagraph> Paragraphs => _paragraphs.AsReadOnly();

    /// <summary>Top-level document blocks in order.</summary>
    public IReadOnlyList<IRtfBlock> Blocks => _blocks.AsReadOnly();

    /// <summary>Font table entries.</summary>
    public IReadOnlyList<RtfFont> Fonts => _fonts.AsReadOnly();

    /// <summary>Color table entries. Index zero is the RTF auto/default color slot.</summary>
    public IReadOnlyList<RtfColor> Colors => _colors.AsReadOnly();

    /// <summary>Stylesheet entries.</summary>
    public IReadOnlyList<RtfStyle> Styles => _styles.AsReadOnly();

    /// <summary>List table definitions.</summary>
    public IReadOnlyList<RtfListDefinition> ListDefinitions => _listDefinitions.AsReadOnly();

    /// <summary>List override table entries.</summary>
    public IReadOnlyList<RtfListOverride> ListOverrides => _listOverrides.AsReadOnly();

    /// <summary>Document headers and footers in write order.</summary>
    public IReadOnlyList<RtfHeaderFooter> HeaderFooters => _headerFooters.AsReadOnly();

    /// <summary>Footnotes, endnotes, and annotations in read or creation order.</summary>
    public IReadOnlyList<RtfNote> Notes => _notes.AsReadOnly();

    /// <summary>Semantic sections in document order when section structure is present.</summary>
    public IReadOnlyList<RtfSection> Sections => _sections.AsReadOnly();

    /// <summary>Custom document properties from the RTF user properties table.</summary>
    public IReadOnlyList<RtfUserProperty> UserProperties => _userProperties.AsReadOnly();

    /// <summary>Document variables from RTF document variable destinations.</summary>
    public IReadOnlyList<RtfDocumentVariable> DocumentVariables => _documentVariables.AsReadOnly();

    /// <summary>Revision author table entries used by run-level revision metadata.</summary>
    public IReadOnlyList<RtfRevisionAuthor> RevisionAuthors => _revisionAuthors.AsReadOnly();

    /// <summary>Revision save identifiers from the RTF <c>rsidtbl</c> destination.</summary>
    public IReadOnlyList<int> RevisionSaveIds => _revisionSaveIds.AsReadOnly();

    /// <summary>Root revision save identifier from <c>\rsidroot</c>.</summary>
    public int? RevisionRootSaveId { get; set; }

    /// <summary>File table references from the document header.</summary>
    public IReadOnlyList<RtfFileReference> FileReferences => _fileReferences.AsReadOnly();

    /// <summary>XML namespace declarations from the RTF <c>xmlnstbl</c> destination.</summary>
    public IReadOnlyList<RtfXmlNamespace> XmlNamespaces => _xmlNamespaces.AsReadOnly();

    /// <summary>Document information metadata.</summary>
    public RtfDocumentInfo Info { get; } = new RtfDocumentInfo();

    /// <summary>Document page size and margins.</summary>
    public RtfPageSetup PageSetup { get; } = new RtfPageSetup();

    /// <summary>Document-level settings such as view, protection, and default tabs.</summary>
    public RtfDocumentSettings Settings { get; } = new RtfDocumentSettings();

    /// <summary>Document-level footnote and endnote numbering settings.</summary>
    public RtfNoteSettings NoteSettings { get; } = new RtfNoteSettings();

    /// <summary>Creates an empty RTF document.</summary>
    public static RtfDocument Create() {
        var document = new RtfDocument();
        document._fonts.Add(new RtfFont(0, "Calibri"));
        return document;
    }

    /// <summary>Reads RTF from a string.</summary>
    public static RtfReadResult Read(string rtf, RtfReadOptions? options = null) {
        return Read(rtf, options, CancellationToken.None);
    }

    /// <summary>Reads RTF from a string with cooperative cancellation during CPU-bound parsing.</summary>
    public static RtfReadResult Read(string rtf, RtfReadOptions? options, CancellationToken cancellationToken) {
        if (rtf == null) throw new ArgumentNullException(nameof(rtf));
        RtfReadOptions readOptions = options ?? RtfReadOptions.CreateOfficeIMOProfile();
        RtfSyntaxTree tree = RtfSyntaxTree.Parse(rtf, readOptions, cancellationToken);
        return RtfSemanticReader.Read(tree, readOptions, cancellationToken);
    }

    /// <summary>Loads RTF from a file.</summary>
    public static RtfReadResult Load(string path, RtfReadOptions? options = null, Encoding? encoding = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        RtfReadOptions readOptions = options ?? RtfReadOptions.CreateOfficeIMOProfile();
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Load(stream, readOptions, encoding);
    }

    /// <summary>Loads RTF from source bytes using the byte-preserving lossless representation.</summary>
    public static RtfReadResult Load(byte[] bytes, RtfReadOptions? options = null) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        RtfReadOptions readOptions = options ?? RtfReadOptions.CreateOfficeIMOProfile();
        new RtfReadLimitGuard(readOptions, CancellationToken.None).CheckInputBytes(bytes.LongLength);
        return Read(RtfBytePreservingEncoding.GetString(bytes), readOptions);
    }

    /// <summary>Loads RTF from a stream.</summary>
    public static RtfReadResult Load(Stream stream, RtfReadOptions? options = null, Encoding? encoding = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        RtfReadOptions readOptions = options ?? RtfReadOptions.CreateOfficeIMOProfile();
        byte[] bytes = RtfBytePreservingEncoding.ReadBytesToEnd(stream, readOptions.MaxInputBytes, CancellationToken.None);
        string rtf = DecodeInput(bytes, encoding);
        return Read(rtf, readOptions);
    }

    private static string DecodeInput(byte[] bytes, Encoding? encoding) {
        if (encoding == null) return RtfBytePreservingEncoding.GetString(bytes);
        using var memory = new MemoryStream(bytes, writable: false);
        using var reader = new StreamReader(memory, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: false);
        return reader.ReadToEnd();
    }

    /// <summary>Adds a paragraph to the document.</summary>
    public RtfParagraph AddParagraph(string? text = null) {
        var paragraph = new RtfParagraph();
        if (!string.IsNullOrEmpty(text)) {
            paragraph.AddText(text!);
        }

        _paragraphs.Add(paragraph);
        _blocks.Add(paragraph);
        return paragraph;
    }

    /// <summary>Adds a semantic section to the document.</summary>
    public RtfSection AddSection(RtfSectionBreakKind breakKind = RtfSectionBreakKind.NextPage) {
        var section = new RtfSection(this) {
            BreakKind = breakKind
        };
        _sections.Add(section);
        return section;
    }

    /// <summary>Adds a table block to the document.</summary>
    public RtfTable AddTable(int rows, int columns) {
        if (rows < 0) throw new ArgumentOutOfRangeException(nameof(rows), "Row count cannot be negative.");
        if (columns <= 0) throw new ArgumentOutOfRangeException(nameof(columns), "Column count must be greater than zero.");

        var table = new RtfTable();
        const int defaultColumnWidthTwips = 2400;
        for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
            RtfTableRow row = table.AddRow();
            for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
                row.AddCell((columnIndex + 1) * defaultColumnWidthTwips);
            }
        }

        _blocks.Add(table);
        return table;
    }

    /// <summary>Adds a picture block to the document.</summary>
    public RtfImage AddImage(RtfImageFormat format, byte[] data) {
        var image = new RtfImage(format, data);
        _blocks.Add(image);
        return image;
    }

    /// <summary>Adds an embedded or linked object block to the document.</summary>
    public RtfObject AddObject(RtfObjectKind kind = RtfObjectKind.Unknown, byte[]? data = null) {
        var rtfObject = new RtfObject(kind, data);
        _blocks.Add(rtfObject);
        return rtfObject;
    }

    /// <summary>Adds a drawing shape block to the document.</summary>
    public RtfShape AddShape() {
        var shape = new RtfShape();
        _blocks.Add(shape);
        return shape;
    }

    /// <summary>Adds a header or footer destination.</summary>
    public RtfHeaderFooter AddHeaderFooter(RtfHeaderFooterKind kind) {
        var headerFooter = new RtfHeaderFooter(kind);
        _headerFooters.Add(headerFooter);
        return headerFooter;
    }

    /// <summary>Adds a header destination.</summary>
    public RtfHeaderFooter AddHeader(RtfHeaderFooterKind kind = RtfHeaderFooterKind.Header) {
        if (kind != RtfHeaderFooterKind.Header && kind != RtfHeaderFooterKind.LeftHeader &&
            kind != RtfHeaderFooterKind.RightHeader && kind != RtfHeaderFooterKind.FirstHeader) {
            throw new ArgumentException("Header kind must be a header destination.", nameof(kind));
        }

        return AddHeaderFooter(kind);
    }

    /// <summary>Adds a footer destination.</summary>
    public RtfHeaderFooter AddFooter(RtfHeaderFooterKind kind = RtfHeaderFooterKind.Footer) {
        if (kind != RtfHeaderFooterKind.Footer && kind != RtfHeaderFooterKind.LeftFooter &&
            kind != RtfHeaderFooterKind.RightFooter && kind != RtfHeaderFooterKind.FirstFooter) {
            throw new ArgumentException("Footer kind must be a footer destination.", nameof(kind));
        }

        return AddHeaderFooter(kind);
    }

    /// <summary>Adds a detached footnote or annotation.</summary>
    public RtfNote AddNote(RtfNoteKind kind) {
        var note = new RtfNote(kind);
        _notes.Add(note);
        return note;
    }

    /// <summary>Adds or returns a list definition.</summary>
    public RtfListDefinition AddListDefinition(int id, string? name = null) {
        RtfListDefinition? existing = _listDefinitions.FirstOrDefault(list => list.Id == id);
        if (existing != null) {
            existing.Name = name ?? existing.Name;
            return existing;
        }

        var definition = new RtfListDefinition(id) {
            Name = name
        };
        _listDefinitions.Add(definition);
        return definition;
    }

    /// <summary>Adds or returns a list override.</summary>
    public RtfListOverride AddListOverride(int id, int listId) {
        RtfListOverride? existing = _listOverrides.FirstOrDefault(item => item.Id == id);
        if (existing != null) return existing;

        var listOverride = new RtfListOverride(id, listId);
        _listOverrides.Add(listOverride);
        return listOverride;
    }

    /// <summary>Adds a font to the font table or returns the existing id.</summary>
    public int AddFont(string name) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Font name cannot be empty.", nameof(name));
        RtfFont? existing = _fonts.FirstOrDefault(font => string.Equals(font.Name, name, StringComparison.OrdinalIgnoreCase));
        if (existing != null) return existing.Id;

        int id = _fonts.Count == 0 ? 0 : _fonts.Max(font => font.Id) + 1;
        _fonts.Add(new RtfFont(id, name));
        return id;
    }

    /// <summary>Adds a color to the color table and returns its one-based RTF color index.</summary>
    public int AddColor(byte red, byte green, byte blue) {
        _colors.Add(new RtfColor(red, green, blue));
        return _colors.Count;
    }

    /// <summary>Adds or replaces a stylesheet entry.</summary>
    public RtfStyle AddStyle(int id, string name, RtfStyleKind kind = RtfStyleKind.Paragraph) {
        RtfStyle? existing = _styles.FirstOrDefault(style => style.Id == id && style.Kind == kind);
        if (existing != null) {
            existing.Name = name;
            return existing;
        }

        var style = new RtfStyle(id, name, kind);
        _styles.Add(style);
        return style;
    }

    /// <summary>Adds a custom document property.</summary>
    public RtfUserProperty AddUserProperty(string name, int? typeCode = null, string? staticValue = null) {
        var property = new RtfUserProperty(name, typeCode, staticValue);
        _userProperties.Add(property);
        return property;
    }

    /// <summary>Adds an existing custom document property.</summary>
    public RtfUserProperty AddUserProperty(RtfUserProperty property) {
        if (property == null) throw new ArgumentNullException(nameof(property));
        _userProperties.Add(property);
        return property;
    }

    /// <summary>Adds a document variable.</summary>
    public RtfDocumentVariable AddDocumentVariable(string name, string value) {
        var variable = new RtfDocumentVariable(name, value);
        _documentVariables.Add(variable);
        return variable;
    }

    /// <summary>Adds a revision author and returns its zero-based table index.</summary>
    public int AddRevisionAuthor(string name) {
        _revisionAuthors.Add(new RtfRevisionAuthor(name));
        return _revisionAuthors.Count - 1;
    }

    /// <summary>Adds a revision save identifier to the document <c>rsidtbl</c>.</summary>
    public RtfDocument AddRevisionSaveId(int id) {
        if (id < 0) throw new ArgumentOutOfRangeException(nameof(id), "Revision save id cannot be negative.");
        _revisionSaveIds.Add(id);
        return this;
    }

    /// <summary>Sets the root revision save identifier represented by <c>\rsidroot</c>.</summary>
    public RtfDocument SetRevisionRootSaveId(int? id) {
        if (id.HasValue && id.Value < 0) throw new ArgumentOutOfRangeException(nameof(id), "Revision root save id cannot be negative.");
        RevisionRootSaveId = id;
        return this;
    }

    /// <summary>Adds a file table reference and returns it.</summary>
    public RtfFileReference AddFileReference(string path, RtfFileSource sources = RtfFileSource.Ntfs) {
        if (path == null) throw new ArgumentNullException(nameof(path));

        int id = _fileReferences.Count == 0 ? 0 : _fileReferences.Max(file => file.Id) + 1;
        var reference = new RtfFileReference(id, path) {
            Sources = sources
        };
        _fileReferences.Add(reference);
        return reference;
    }

    /// <summary>Adds an XML namespace declaration to the document namespace table.</summary>
    public RtfXmlNamespace AddXmlNamespace(int id, string uri) {
        if (string.IsNullOrWhiteSpace(uri)) throw new ArgumentException("XML namespace URI cannot be empty.", nameof(uri));

        RtfXmlNamespace? existing = _xmlNamespaces.FirstOrDefault(ns => ns.Id == id);
        if (existing != null) {
            existing.Uri = uri;
            return existing;
        }

        var xmlNamespace = new RtfXmlNamespace(id, uri);
        _xmlNamespaces.Add(xmlNamespace);
        return xmlNamespace;
    }

    /// <summary>Serializes the document to RTF.</summary>
    public string ToRtf(RtfWriteOptions? options = null) => RtfDocumentWriter.Write(this, options ?? new RtfWriteOptions());

    /// <summary>Serializes the document to encoded RTF bytes.</summary>
    public byte[] ToBytes(RtfWriteOptions? options = null, Encoding? encoding = null) {
        return (encoding ?? Encoding.UTF8).GetBytes(ToRtf(options));
    }

    /// <summary>Serializes the document to an encoded RTF memory stream.</summary>
    public MemoryStream ToMemoryStream(RtfWriteOptions? options = null, Encoding? encoding = null) {
        return new MemoryStream(ToBytes(options, encoding), writable: false);
    }

    /// <summary>Saves the document to an RTF file.</summary>
    public void Save(string path, RtfWriteOptions? options = null, Encoding? encoding = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        File.WriteAllText(path, ToRtf(options), encoding ?? Encoding.UTF8);
    }

    /// <summary>Saves the document to an RTF stream without closing the stream.</summary>
    public void Save(Stream stream, RtfWriteOptions? options = null, Encoding? encoding = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        byte[] bytes = ToBytes(options, encoding);
        stream.Write(bytes, 0, bytes.Length);
    }

    internal void AddParsedParagraph(RtfParagraph paragraph) {
        paragraph = paragraph ?? throw new ArgumentNullException(nameof(paragraph));
        _paragraphs.Add(paragraph);
        _blocks.Add(paragraph);
    }

    internal void AddParsedBlock(IRtfBlock block) {
        if (block == null) throw new ArgumentNullException(nameof(block));
        _blocks.Add(block);
        if (block is RtfParagraph paragraph) {
            _paragraphs.Add(paragraph);
        }
    }

    internal void AddParsedSection(RtfSection section) {
        _sections.Add(section ?? throw new ArgumentNullException(nameof(section)));
    }

    internal void ReplaceFonts(IEnumerable<RtfFont> fonts) {
        _fonts.Clear();
        _fonts.AddRange(fonts ?? Array.Empty<RtfFont>());
        if (_fonts.Count == 0) {
            _fonts.Add(new RtfFont(0, "Calibri"));
        }
    }

    internal void ReplaceColors(IEnumerable<RtfColor> colors) {
        _colors.Clear();
        _colors.AddRange(colors ?? Array.Empty<RtfColor>());
    }

    internal void ReplaceStyles(IEnumerable<RtfStyle> styles) {
        _styles.Clear();
        _styles.AddRange(styles ?? Array.Empty<RtfStyle>());
    }

    internal void ReplaceListDefinitions(IEnumerable<RtfListDefinition> listDefinitions) {
        _listDefinitions.Clear();
        _listDefinitions.AddRange(listDefinitions ?? Array.Empty<RtfListDefinition>());
    }

    internal void ReplaceListOverrides(IEnumerable<RtfListOverride> listOverrides) {
        _listOverrides.Clear();
        _listOverrides.AddRange(listOverrides ?? Array.Empty<RtfListOverride>());
    }

    internal void ReplaceUserProperties(IEnumerable<RtfUserProperty> userProperties) {
        _userProperties.Clear();
        _userProperties.AddRange(userProperties ?? Array.Empty<RtfUserProperty>());
    }

    internal void ReplaceDocumentVariables(IEnumerable<RtfDocumentVariable> documentVariables) {
        _documentVariables.Clear();
        _documentVariables.AddRange(documentVariables ?? Array.Empty<RtfDocumentVariable>());
    }

    internal void ReplaceRevisionAuthors(IEnumerable<RtfRevisionAuthor> revisionAuthors) {
        _revisionAuthors.Clear();
        _revisionAuthors.AddRange(revisionAuthors ?? Array.Empty<RtfRevisionAuthor>());
    }

    internal void ReplaceRevisionSaveIds(IEnumerable<int> revisionSaveIds) {
        _revisionSaveIds.Clear();
        _revisionSaveIds.AddRange(revisionSaveIds ?? Array.Empty<int>());
    }

    internal void ReplaceFileReferences(IEnumerable<RtfFileReference> fileReferences) {
        _fileReferences.Clear();
        _fileReferences.AddRange(fileReferences ?? Array.Empty<RtfFileReference>());
    }

    internal void ReplaceXmlNamespaces(IEnumerable<RtfXmlNamespace> xmlNamespaces) {
        _xmlNamespaces.Clear();
        _xmlNamespaces.AddRange(xmlNamespaces ?? Array.Empty<RtfXmlNamespace>());
    }

    internal void AddParsedHeaderFooter(RtfHeaderFooter headerFooter) {
        _headerFooters.Add(headerFooter ?? throw new ArgumentNullException(nameof(headerFooter)));
    }

    internal void AddParsedNote(RtfNote note) {
        _notes.Add(note ?? throw new ArgumentNullException(nameof(note)));
    }
}
