using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.IWork.Internal;

/// <summary>Validated iWork '09 XML snapshot. The XML stays private to generation-specific projections.</summary>
internal sealed class IWorkLegacyXmlSource {
    private const string Version = "92008102400";
    private const int MaximumElements = 100_000;
    private const int MaximumAttributes = 500_000;
    private const int MaximumDepth = 64;
    private const long MaximumXmlCharacters = 16L * 1024 * 1024;

    internal static readonly XNamespace Sf = "http://developer.apple.com/namespaces/sf";
    internal static readonly XNamespace Sfa = "http://developer.apple.com/namespaces/sfa";
    internal static readonly XNamespace Pages = "http://developer.apple.com/namespaces/sl";
    internal static readonly XNamespace Numbers = "http://developer.apple.com/namespaces/ls";
    internal static readonly XNamespace Keynote = "http://developer.apple.com/namespaces/keynote2";

    private IWorkLegacyXmlSource(IWorkDocumentKind kind, string entryPath, XElement root) {
        Kind = kind;
        EntryPath = entryPath;
        Root = root;
    }

    internal IWorkDocumentKind Kind { get; }
    internal string EntryPath { get; }
    internal XElement Root { get; }

    internal static bool IsIndexPath(string path) =>
        path.Equals("index.xml", StringComparison.OrdinalIgnoreCase)
        || path.Equals("index.xml.gz", StringComparison.OrdinalIgnoreCase)
        || path.Equals("index.apxl", StringComparison.OrdinalIgnoreCase)
        || path.Equals("index.apxl.gz", StringComparison.OrdinalIgnoreCase);

    internal static IWorkLegacyXmlSource Read(IWorkPackageEntry entry, IWorkLoadOptions options,
        CancellationToken cancellationToken) {
        try {
            using XmlReader reader = OpenReader(new MemoryStream(entry.Bytes, writable: false),
                entry.Path, options.MaxEntryBytes, cancellationToken);
            IWorkDocumentKind kind = ReadKind(reader, entry.Path)
                ?? throw new NotSupportedException("The legacy iWork XML index is outside the supported '09 profile.");
            cancellationToken.ThrowIfCancellationRequested();
            XElement root = XElement.Load(reader, LoadOptions.PreserveWhitespace);
            cancellationToken.ThrowIfCancellationRequested();
            return new IWorkLegacyXmlSource(kind, entry.Path, root);
        } catch (XmlException exception) {
            throw new InvalidDataException("The iWork '09 XML index is malformed or exceeds the XML character limit.",
                exception);
        }
    }

    internal static bool Probe(Stream stream, string path, IWorkDocumentKind expectedKind, long maximumEntryBytes,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        using XmlReader reader = OpenReader(stream, path, maximumEntryBytes, cancellationToken);
        return ReadKind(reader, path) == expectedKind;
    }

    private static XmlReader OpenReader(Stream stream, string path, long maximumEntryBytes,
        CancellationToken cancellationToken) {
        Stream content = path.EndsWith(".gz", StringComparison.OrdinalIgnoreCase)
            ? new GZipStream(stream, CompressionMode.Decompress)
            : stream;
        try {
            XmlReader reader = XmlReader.Create(content, new XmlReaderSettings {
                CloseInput = true,
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = Math.Min(MaximumXmlCharacters, maximumEntryBytes),
                MaxCharactersFromEntities = 0
            });
            return new OfficeXmlLimitingReader(reader, "iWork '09 XML index", MaximumDepth,
                MaximumElements, MaximumAttributes, cancellationToken);
        } catch {
            content.Dispose();
            throw;
        }
    }

    private static IWorkDocumentKind? ReadKind(XmlReader reader, string path) {
        if (reader.MoveToContent() != XmlNodeType.Element
            || reader.GetAttribute("version", reader.NamespaceURI) != Version) return null;
        bool isPresentation = path.StartsWith("index.apxl", StringComparison.OrdinalIgnoreCase);
        if (isPresentation) {
            return reader.NamespaceURI == Keynote.NamespaceName && reader.LocalName == "presentation"
                ? IWorkDocumentKind.Keynote : null;
        }
        if (reader.LocalName != "document") return null;
        if (reader.NamespaceURI == Pages.NamespaceName) return IWorkDocumentKind.Pages;
        return reader.NamespaceURI == Numbers.NamespaceName ? IWorkDocumentKind.Numbers : null;
    }
}
