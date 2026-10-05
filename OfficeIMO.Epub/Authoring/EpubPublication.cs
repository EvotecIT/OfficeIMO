using OfficeIMO.Core.Internal;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Epub;

/// <summary>
/// A writable EPUB package retaining unknown XML and entry payloads. Instances are mutable and not thread-safe.
/// Use <see cref="Read"/> to project the current publication through the existing extraction owner.
/// </summary>
public sealed partial class EpubPublication {
    internal static readonly XNamespace Opf = "http://www.idpf.org/2007/opf";
    internal static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
    internal static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    internal static readonly XNamespace Ops = "http://www.idpf.org/2007/ops";
    private readonly XDocument _package;
    private readonly Dictionary<string, EpubManifestItem> _originalManifest;
    private readonly Dictionary<string, byte[]> _entries;
    private readonly Dictionary<string, byte[]> _originalEntries;
    private readonly byte[]? _originalBytes;
    private readonly bool _originalHasZipSignature;
    private readonly int _originalEntryCount;
    private readonly long _originalExpandedBytes;
    private readonly string[] _rootfilePaths;
    private readonly IReadOnlyList<EpubEncryptionInfo> _encryption;
    private readonly string? _originalIdentifier;
    private bool _changed;
    private DateTimeOffset _modifiedAt;
    private readonly long _maximumRetainedBytes;
    private readonly long _maximumEntryBytes;
    private readonly long _maximumMetadataBytes;
    private readonly int _maximumEntries;
    private long _retainedBytes;

    private EpubPublication(string path, XDocument package, Dictionary<string, byte[]> entries,
        byte[]? originalBytes = null, IReadOnlyList<EpubEncryptionInfo>? encryption = null, EpubPublicationLoadOptions? limits = null,
        IReadOnlyList<EpubRootfile>? rootfiles = null, int originalEntryCount = 0, long originalExpandedBytes = 0) {
        PackagePath = path;
        _package = package;
        _originalManifest = package.Root!.Element(Opf + "manifest")?.Elements(Opf + "item")
            .GroupBy(element => (string?)element.Attribute("id") ?? string.Empty, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => new EpubManifestItem(new XElement(group.First()), path), StringComparer.Ordinal)
            ?? new Dictionary<string, EpubManifestItem>(StringComparer.Ordinal);
        _entries = entries;
        _originalEntries = new Dictionary<string, byte[]>(entries, StringComparer.Ordinal);
        _entryOrigins = entries.Keys.ToDictionary(path => path, path => path, StringComparer.Ordinal);
        _originalBytes = originalBytes;
        _originalEntryCount = originalEntryCount;
        _originalExpandedBytes = originalExpandedBytes;
        _rootfilePaths = rootfiles?.Select(rootfile => rootfile.FullPath).ToArray() ?? new[] { path };
        _encryption = encryption ?? Array.Empty<EpubEncryptionInfo>();
        limits ??= new EpubPublicationLoadOptions();
        _originalHasZipSignature = originalBytes != null && OfficeIMO.Provenance.OfficeProvenanceZip.HasCentralDirectorySignature(originalBytes, limits.MaxEntries);
        _maximumRetainedBytes = limits.MaxExpandedBytes;
        _maximumEntryBytes = limits.MaxEntryBytes;
        _maximumMetadataBytes = limits.MaxMetadataBytes;
        _maximumEntries = limits.MaxEntries;
        EnsureEntryBudget();
        _retainedBytes = entries.Values.Sum(data => data.LongLength);
        if (Root.Name != Opf + "package" || (PackageVersion != "3.0" && PackageVersion != "2.0")) {
            throw new NotSupportedException("Writing supports OPF package versions 2.0 and 3.0.");
        }
        RequireSection("metadata"); RequireSection("manifest"); RequireSection("spine");
        _originalIdentifier = Identifier;
        _modifiedAt = DateTimeOffset.UtcNow;
        _changed = originalBytes == null;
        if (originalBytes == null) EnsurePackageBudget(_package);
        _package.Changed += (_, _) => MarkChanged();
    }

    /// <summary>Creates a publication with stable identity; newly authored packages need at least one linear chapter.</summary>
    public static EpubPublication Create(string title, string language = "en", string? identifier = null,
        EpubVersion version = EpubVersion.Epub3, EpubPublicationLoadOptions? retentionLimits = null) {
        EpubPublicationLoadOptions limits = SnapshotLoadOptions(retentionLimits);
        RequireText(title, nameof(title)); RequireText(language, nameof(language));
        EpubLanguageTag.Require(language, nameof(language));
        if (version != EpubVersion.Epub2 && version != EpubVersion.Epub3) throw new ArgumentOutOfRangeException(nameof(version));
        identifier ??= "urn:uuid:" + Guid.NewGuid().ToString("D");
        RequireText(identifier, nameof(identifier));
        var package = new XDocument(new XElement(Opf + "package",
            new XAttribute(XNamespace.Xml + "lang", language),
            new XAttribute("version", version == EpubVersion.Epub3 ? "3.0" : "2.0"),
            new XAttribute("unique-identifier", "publication-id"),
            new XElement(Opf + "metadata", new XAttribute(XNamespace.Xmlns + "dc", Dc.NamespaceName),
                new XElement(Dc + "identifier", new XAttribute("id", "publication-id"), identifier),
                new XElement(Dc + "title", title), new XElement(Dc + "language", language)),
            new XElement(Opf + "manifest"), new XElement(Opf + "spine")));
        const string path = "EPUB/package.opf";
        XNamespace container = "urn:oasis:names:tc:opendocument:xmlns:container";
        var entries = new Dictionary<string, byte[]>(StringComparer.Ordinal) {
            ["mimetype"] = Encoding.ASCII.GetBytes("application/epub+zip"),
            ["META-INF/container.xml"] = SerializeXml(new XDocument(new XElement(container + "container",
                new XAttribute("version", "1.0"), new XElement(container + "rootfiles",
                    new XElement(container + "rootfile", new XAttribute("full-path", path),
                        new XAttribute("media-type", "application/oebps-package+xml"))))))
        };
        var publication = new EpubPublication(path, package, entries, limits: limits);
        publication.InitializeNavigation();
        return publication;
    }

    /// <summary>Loads all bounded entries for editing without conflating retained package content with extracted chapters.</summary>
    public static EpubPublication Load(Stream stream, EpubPublicationLoadOptions? options = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        EpubPublicationLoadOptions effective = SnapshotLoadOptions(options);
        byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, cancellationToken, effective.MaxInputBytes);
        return LoadBytes(bytes, effective, cancellationToken);
    }

    /// <summary>Loads a file for bounded package editing.</summary>
    public static EpubPublication Load(string path, EpubPublicationLoadOptions? options = null, CancellationToken cancellationToken = default) {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Load(stream, options, cancellationToken);
    }

    /// <summary>Asynchronously loads a caller-owned stream for editing.</summary>
    public static async Task<EpubPublication> LoadAsync(Stream stream, EpubPublicationLoadOptions? options = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        EpubPublicationLoadOptions effective = SnapshotLoadOptions(options);
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, effective.MaxInputBytes).ConfigureAwait(false);
        return LoadBytes(bytes, effective, cancellationToken);
    }

    /// <summary>Asynchronously loads a file for editing.</summary>
    public static async Task<EpubPublication> LoadAsync(string path, EpubPublicationLoadOptions? options = null, CancellationToken cancellationToken = default) {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, true);
        return await LoadAsync(stream, options, cancellationToken).ConfigureAwait(false);
    }

    private static EpubPublication LoadBytes(byte[] bytes, EpubPublicationLoadOptions options, CancellationToken token) {
        var source = EpubReader.ReadEditablePackage(bytes, options, token);
        return new EpubPublication(source.OpfPath, ParseXml(source.Entries[source.OpfPath], options.MaxMetadataBytes),
            source.Entries, bytes, source.Encryption, options, source.Rootfiles, source.EntryCount, source.ExpandedBytes);
    }

    /// <summary>Selected package document path. Other rootfiles and their payloads are preserved.</summary>
    public string PackagePath { get; }
    /// <summary>Declared package version, preserved when editing.</summary>
    public string PackageVersion => (string?)Root.Attribute("version") ?? string.Empty;
    /// <summary>Typed manifest declarations in package order.</summary>
    public IReadOnlyList<EpubManifestItem> Manifest => Array.AsReadOnly(RequireSection("manifest").Elements(Opf + "item")
        .Select(item => new EpubManifestItem(item, PackagePath, this)).ToArray());
    /// <summary>Typed reading positions, including repeated resource references.</summary>
    public IReadOnlyList<EpubSpineItem> Spine => Array.AsReadOnly(RequireSection("spine").Elements(Opf + "itemref")
        .Select(item => new EpubSpineItem(item, this)).ToArray());
    /// <summary>All retained entry paths, including unmanifested extension payloads.</summary>
    public IReadOnlyList<string> EntryPaths => Array.AsReadOnly(_entries.Keys.OrderBy(path => path, StringComparer.Ordinal).ToArray());
    /// <summary>Returns an independent copy of package XML for inspection.</summary>
    public XDocument GetPackageXml() => new XDocument(_package);
    /// <summary>Projects current authored content through the existing EPUB reader, without executing embedded content.</summary>
    public EpubDocument Read(EpubReadOptions? options = null, CancellationToken cancellationToken = default) =>
        EpubDocument.Load(new MemoryStream(Write(cancellationToken: cancellationToken).Bytes, false), options, cancellationToken);

    private XElement Root => _package.Root ?? throw new InvalidDataException("Package XML has no root.");
    private XElement RequireSection(string name) => Root.Element(Opf + name) ?? throw new InvalidDataException("Package section missing: " + name);
    private void EnsureEntryBudget(int additionalEntries = 0) {
        int current = _entries.Count + (_entries.ContainsKey(PackagePath) ? 0 : 1);
        if (additionalEntries > _maximumEntries - current)
            throw new InvalidDataException("Publication exceeds the retained entry-count limit.");
    }
    private void MarkChanged() {
        if (!_changed) _modifiedAt = DateTimeOffset.UtcNow;
        _changed = true;
    }
    private static void RequireText(string value, string name) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Value cannot be empty.", name);
        XmlConvert.VerifyXmlChars(value);
    }
    private static EpubPublicationLoadOptions SnapshotLoadOptions(EpubPublicationLoadOptions? options) {
        options ??= new EpubPublicationLoadOptions();
        var result = new EpubPublicationLoadOptions { MaxInputBytes = options.MaxInputBytes, MaxExpandedBytes = options.MaxExpandedBytes,
            MaxEntryBytes = options.MaxEntryBytes, MaxMetadataBytes = options.MaxMetadataBytes, MaxEntries = options.MaxEntries };
        if (result.MaxInputBytes < 1 || result.MaxExpandedBytes < 1 || result.MaxEntryBytes < 1 || result.MaxMetadataBytes < 1 || result.MaxEntries < 1)
            throw new ArgumentOutOfRangeException(nameof(options), "All input bounds must be positive.");
        return result;
    }
    internal static XDocument ParseXml(byte[] data, long maximumBytes) {
        if (data.LongLength > maximumBytes) throw new InvalidDataException("XML exceeds the configured byte limit.");
        using var stream = new MemoryStream(data, false);
        using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Ignore, XmlResolver = null, MaxCharactersInDocument = maximumBytes
        });
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }
    internal static byte[] SerializeXml(XDocument document, long maximumBytes = 128L * 1024 * 1024) {
        using var output = new OfficeBoundedMemoryStream(maximumBytes);
        using (XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings {
            Encoding = new UTF8Encoding(false, true), Indent = false, NewLineHandling = NewLineHandling.Entitize
        })) document.Save(writer);
        return output.ToArray();
    }
}
