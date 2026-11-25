using System.Text;

namespace OfficeIMO.Core.Internal.OpenXml;

/// <summary>
/// Namespace identifiers for the Open XML pull reader. Well-known namespaces have fixed ids; ISO/IEC 29500 Strict
/// URIs map to the same id as their transitional counterpart, so readers match one id whichever conformance class
/// the part uses. Other URIs get ids as they are first declared in a part.
/// </summary>
internal static class OfficeXmlNamespaces {
    internal const int None = 0;
    internal const int Xml = 1;
    internal const int Xmlns = 2;
    internal const int MarkupCompatibility = 3;
    internal const int ContentTypes = 4;
    internal const int PackageRelationships = 5;
    internal const int Relationships = 6;
    internal const int WordprocessingML = 7;
    internal const int SpreadsheetML = 8;
    internal const int PresentationML = 9;
    internal const int DrawingML = 10;
    internal const int Math = 11;
    internal const int WordprocessingDrawing = 12;
    internal const int SpreadsheetDrawing = 13;
    internal const int Picture = 14;
    internal const int Chart = 15;
    internal const int ChartDrawing = 16;
    internal const int Diagram = 17;
    internal const int LockedCanvas = 18;
    internal const int DrawingCompatibility = 19;
    internal const int SharedTypes = 20;
    internal const int CustomXml = 21;
    internal const int Bibliography = 22;
    internal const int SchemaLibrary = 23;
    internal const int ExtendedProperties = 24;
    internal const int CustomProperties = 25;
    internal const int DocPropsVTypes = 26;
    internal const int Vml = 27;
    internal const int VmlOffice = 28;
    internal const int VmlWord = 29;
    internal const int VmlExcel = 30;
    internal const int VmlPowerPoint = 31;
    internal const int CoreProperties = 32;
    internal const int DublinCore = 33;
    internal const int DublinCoreTerms = 34;

    /// <summary>The first id handed to a namespace that is not listed here.</summary>
    internal const int FirstDynamic = 35;

    /// <summary>Transitional relationship-type prefix that Strict relationship types are rewritten to.</summary>
    internal const string TransitionalRelationshipTypePrefix = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string StrictRelationshipTypePrefix = "http://purl.oclc.org/ooxml/officeDocument/relationships/";

    // Namespaces ECMA-376 defines (Part 1 for markup, Part 2 for packaging, Part 3 for markup compatibility, Part 4
    // for VML) plus the Dublin Core vocabularies of the core-properties part. Strict URIs come from ISO/IEC 29500-1.
    private static readonly KeyValuePair<string, int>[] KnownUris = {
        new("http://www.w3.org/XML/1998/namespace", Xml),
        new("http://www.w3.org/2000/xmlns/", Xmlns),
        new("http://schemas.openxmlformats.org/markup-compatibility/2006", MarkupCompatibility),
        new("http://schemas.openxmlformats.org/package/2006/content-types", ContentTypes),
        new("http://schemas.openxmlformats.org/package/2006/relationships", PackageRelationships),
        new("http://schemas.openxmlformats.org/package/2006/metadata/core-properties", CoreProperties),
        new("http://purl.org/dc/elements/1.1/", DublinCore),
        new("http://purl.org/dc/terms/", DublinCoreTerms),
        new("http://schemas.openxmlformats.org/officeDocument/2006/relationships", Relationships),
        new("http://purl.oclc.org/ooxml/officeDocument/relationships", Relationships),
        new("http://schemas.openxmlformats.org/wordprocessingml/2006/main", WordprocessingML),
        new("http://purl.oclc.org/ooxml/wordprocessingml/main", WordprocessingML),
        new("http://schemas.openxmlformats.org/spreadsheetml/2006/main", SpreadsheetML),
        new("http://purl.oclc.org/ooxml/spreadsheetml/main", SpreadsheetML),
        new("http://schemas.openxmlformats.org/presentationml/2006/main", PresentationML),
        new("http://purl.oclc.org/ooxml/presentationml/main", PresentationML),
        new("http://schemas.openxmlformats.org/drawingml/2006/main", DrawingML),
        new("http://purl.oclc.org/ooxml/drawingml/main", DrawingML),
        new("http://schemas.openxmlformats.org/officeDocument/2006/math", Math),
        new("http://purl.oclc.org/ooxml/officeDocument/math", Math),
        new("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", WordprocessingDrawing),
        new("http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing", WordprocessingDrawing),
        new("http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", SpreadsheetDrawing),
        new("http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing", SpreadsheetDrawing),
        new("http://schemas.openxmlformats.org/drawingml/2006/picture", Picture),
        new("http://purl.oclc.org/ooxml/drawingml/picture", Picture),
        new("http://schemas.openxmlformats.org/drawingml/2006/chart", Chart),
        new("http://purl.oclc.org/ooxml/drawingml/chart", Chart),
        new("http://schemas.openxmlformats.org/drawingml/2006/chartDrawing", ChartDrawing),
        new("http://purl.oclc.org/ooxml/drawingml/chartDrawing", ChartDrawing),
        new("http://schemas.openxmlformats.org/drawingml/2006/diagram", Diagram),
        new("http://purl.oclc.org/ooxml/drawingml/diagram", Diagram),
        new("http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas", LockedCanvas),
        new("http://purl.oclc.org/ooxml/drawingml/lockedCanvas", LockedCanvas),
        new("http://schemas.openxmlformats.org/drawingml/2006/compatibility", DrawingCompatibility),
        new("http://purl.oclc.org/ooxml/drawingml/compatibility", DrawingCompatibility),
        new("http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes", SharedTypes),
        new("http://purl.oclc.org/ooxml/officeDocument/sharedTypes", SharedTypes),
        new("http://schemas.openxmlformats.org/officeDocument/2006/customXml", CustomXml),
        new("http://purl.oclc.org/ooxml/officeDocument/customXml", CustomXml),
        new("http://schemas.openxmlformats.org/officeDocument/2006/bibliography", Bibliography),
        new("http://purl.oclc.org/ooxml/officeDocument/bibliography", Bibliography),
        new("http://schemas.openxmlformats.org/schemaLibrary/2006/main", SchemaLibrary),
        new("http://purl.oclc.org/ooxml/schemaLibrary/main", SchemaLibrary),
        new("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", ExtendedProperties),
        new("http://purl.oclc.org/ooxml/officeDocument/extendedProperties", ExtendedProperties),
        new("http://schemas.openxmlformats.org/officeDocument/2006/custom-properties", CustomProperties),
        new("http://purl.oclc.org/ooxml/officeDocument/customProperties", CustomProperties),
        new("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", DocPropsVTypes),
        new("http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes", DocPropsVTypes),
        new("urn:schemas-microsoft-com:vml", Vml),
        new("urn:schemas-microsoft-com:office:office", VmlOffice),
        new("urn:schemas-microsoft-com:office:word", VmlWord),
        new("urn:schemas-microsoft-com:office:excel", VmlExcel),
        new("urn:schemas-microsoft-com:office:powerpoint", VmlPowerPoint),
    };

    private static readonly byte[][] KnownUriBytes = CreateKnownUriBytes();

    private static byte[][] CreateKnownUriBytes() {
        var bytes = new byte[KnownUris.Length][];
        for (int index = 0; index < KnownUris.Length; index++) bytes[index] = Encoding.UTF8.GetBytes(KnownUris[index].Key);
        return bytes;
    }

    /// <summary>
    /// Every fixed id: the namespaces an ECMA-376 first-edition consumer understands. Markup-compatibility processing
    /// treats any other namespace as not understood.
    /// </summary>
    internal static bool IsEcma376(int id) => id > None && id < FirstDynamic;

    /// <summary>Rewrites a Strict relationship type to its transitional spelling; other types are returned unchanged.</summary>
    internal static string NormalizeRelationshipType(string type) =>
        type.StartsWith(StrictRelationshipTypePrefix, StringComparison.Ordinal)
            ? TransitionalRelationshipTypePrefix + type.Substring(StrictRelationshipTypePrefix.Length)
            : type;

    /// <summary>The exact URI id of a namespace's transitional spelling; for a namespace that is not listed, its own id.</summary>
    internal static int TransitionalUriId(int namespaceId) {
        if (namespaceId <= Xmlns && namespaceId != Xml || namespaceId >= FirstDynamic) return namespaceId;
        for (int index = 0; index < KnownUris.Length; index++) {
            if (KnownUris[index].Value == namespaceId) return KnownUriIdBase + index;
        }
        return namespaceId;
    }

    /// <summary>Whether a URI id is a Strict spelling (ISO/IEC 29500-1) of a listed namespace.</summary>
    internal static bool IsStrictUriId(int uriId) =>
        uriId >= KnownUriIdBase && KnownUris[uriId - KnownUriIdBase].Key.StartsWith("http://purl.oclc.org/ooxml/", StringComparison.Ordinal);

    // URI ids of listed namespaces sit far above any dynamic id.
    private const int KnownUriIdBase = 1 << 24;

    /// <summary>Per-package table that assigns ids to namespace URIs declared in the parts it reads and caches local names.</summary>
    internal sealed class Table {
        private const int NameCacheSize = 256;
        private readonly List<byte[]> _dynamicUris = new();
        private string?[]? _names;

        /// <summary>The cache slot for an element local name with this hash; readers of one package share it.</summary>
        internal ref string? NameSlot(int hash) => ref (_names ??= new string?[NameCacheSize])[hash & (NameCacheSize - 1)];
        private int _lastKnown = -1;

        // Only trees intern qualified names; streaming readers of small parts never allocate these.
        private OfficeXmlName[] _xmlNames = Array.Empty<OfficeXmlName>();
        private int _xmlNameCount;
        // Open addressing over name ids plus one; zero marks an empty slot.
        private int[] _xmlNameSlots = Array.Empty<int>();

        /// <summary>
        /// The id of a qualified name as an element or attribute uses it: local name, prefix (null when a writer picks
        /// it), folded namespace and exact URI. Trees of the package share the ids, so nodes store one integer.
        /// </summary>
        internal int GetNameId(string localName, string? prefix, int namespaceId, int uriId) {
            if (_xmlNameSlots.Length == 0) {
                _xmlNames = new OfficeXmlName[64];
                _xmlNameSlots = new int[128];
            }
            int hash = ((localName.GetHashCode() * 31 + (prefix == null ? 1 : prefix.GetHashCode())) * 31 + namespaceId) * 31 + uriId;
            int mask = _xmlNameSlots.Length - 1;
            int slot = hash & mask;
            while (true) {
                int entry = _xmlNameSlots[slot];
                if (entry == 0) break;
                ref OfficeXmlName name = ref _xmlNames[entry - 1];
                if (name.NamespaceId == namespaceId && name.UriId == uriId && string.Equals(name.LocalName, localName, StringComparison.Ordinal)
                    && string.Equals(name.Prefix, prefix, StringComparison.Ordinal)) {
                    return entry - 1;
                }
                slot = (slot + 1) & mask;
            }
            if (_xmlNameCount == _xmlNames.Length) Array.Resize(ref _xmlNames, _xmlNames.Length * 2);
            int id = _xmlNameCount++;
            _xmlNames[id] = new OfficeXmlName(localName, prefix, namespaceId, uriId);
            _xmlNameSlots[slot] = id + 1;
            if (_xmlNameCount * 2 > _xmlNameSlots.Length) RehashNames();
            return id;
        }

        /// <summary>The name behind an id from <see cref="GetNameId"/>.</summary>
        internal ref readonly OfficeXmlName GetName(int id) => ref _xmlNames[id];

        private byte[]?[] _xmlNameUtf8 = Array.Empty<byte[]?>();

        /// <summary>The UTF-8 bytes of a name's local part, for writers.</summary>
        internal byte[] GetLocalNameUtf8(int id) {
            if (id >= _xmlNameUtf8.Length) Array.Resize(ref _xmlNameUtf8, System.Math.Max(System.Math.Max(64, _xmlNameUtf8.Length * 2), id + 1));
            return _xmlNameUtf8[id] ??= Encoding.UTF8.GetBytes(_xmlNames[id].LocalName);
        }

        private void RehashNames() {
            var slots = new int[_xmlNameSlots.Length * 2];
            int mask = slots.Length - 1;
            for (int id = 0; id < _xmlNameCount; id++) {
                ref OfficeXmlName name = ref _xmlNames[id];
                int hash = ((name.LocalName.GetHashCode() * 31 + (name.Prefix == null ? 1 : name.Prefix.GetHashCode())) * 31 + name.NamespaceId) * 31 + name.UriId;
                int slot = hash & mask;
                while (slots[slot] != 0) slot = (slot + 1) & mask;
                slots[slot] = id + 1;
            }
            _xmlNameSlots = slots;
        }

        /// <summary>Returns the id of the UTF-8 URI in <paramref name="data"/>, registering it when it is new.</summary>
        internal int GetId(byte[] data, int start, int length) => GetId(data, start, length, out _);

        /// <summary>Returns the id of a URI and, in <paramref name="uriId"/>, an id unique to its exact spelling, so a
        /// Strict URI and its transitional counterpart share an id but not a URI id.</summary>
        internal int GetId(byte[] data, int start, int length, out int uriId) {
            // Packages repeat the same few declarations on many elements, so try the last match first.
            if (_lastKnown >= 0 && Matches(KnownUriBytes[_lastKnown], data, start, length)) {
                uriId = KnownUriIdBase + _lastKnown;
                return KnownUris[_lastKnown].Value;
            }
            for (int index = 0; index < KnownUriBytes.Length; index++) {
                if (Matches(KnownUriBytes[index], data, start, length)) {
                    _lastKnown = index;
                    uriId = KnownUriIdBase + index;
                    return KnownUris[index].Value;
                }
            }
            // Other URIs are found through a hash index, so a part that declares many namespaces stays linear.
            int hash = HashOf(data, start, length);
            int mask = _dynamicSlots.Length - 1;
            int slot = hash & mask;
            while (_dynamicSlots[slot] != 0) {
                int index = _dynamicSlots[slot] - 1;
                if (Matches(_dynamicUris[index], data, start, length)) return uriId = FirstDynamic + index;
                slot = (slot + 1) & mask;
            }
            var uri = new byte[length];
            Buffer.BlockCopy(data, start, uri, 0, length);
            _dynamicUris.Add(uri);
            _dynamicSlots[slot] = _dynamicUris.Count;
            if (_dynamicUris.Count * 2 > _dynamicSlots.Length) RehashDynamicUris();
            return uriId = FirstDynamic + _dynamicUris.Count - 1;
        }

        /// <summary>Returns the id of a URI given as a string, registering it when it is new.</summary>
        internal int GetId(string uri) => GetId(uri, out _);

        internal int GetId(string uri, out int uriId) {
            for (int index = 0; index < KnownUris.Length; index++) {
                if (string.Equals(KnownUris[index].Key, uri, StringComparison.Ordinal)) {
                    uriId = KnownUriIdBase + index;
                    return KnownUris[index].Value;
                }
            }
            byte[] bytes = Encoding.UTF8.GetBytes(uri);
            return GetId(bytes, 0, bytes.Length, out uriId);
        }

        /// <summary>Returns the exact URI behind a URI id from <see cref="GetId(byte[], int, int, out int)"/>.</summary>
        internal string GetUriById(int uriId) =>
            uriId >= KnownUriIdBase ? KnownUris[uriId - KnownUriIdBase].Key : GetUri(uriId);

        /// <summary>Returns the URI behind an id; for a fixed id this is its transitional spelling.</summary>
        internal string GetUri(int id) {
            if (id >= FirstDynamic) return Encoding.UTF8.GetString(_dynamicUris[id - FirstDynamic]);
            foreach (KeyValuePair<string, int> known in KnownUris) {
                if (known.Value == id) return known.Key;
            }
            return string.Empty;
        }

        private int[] _dynamicSlots = new int[16];

        private void RehashDynamicUris() {
            var slots = new int[_dynamicSlots.Length * 2];
            int mask = slots.Length - 1;
            for (int index = 0; index < _dynamicUris.Count; index++) {
                byte[] uri = _dynamicUris[index];
                int slot = HashOf(uri, 0, uri.Length) & mask;
                while (slots[slot] != 0) slot = (slot + 1) & mask;
                slots[slot] = index + 1;
            }
            _dynamicSlots = slots;
        }

        // FNV-1a over the URI bytes.
        private static int HashOf(byte[] data, int start, int length) {
            uint hash = 2166136261;
            for (int index = start; index < start + length; index++) hash = (hash ^ data[index]) * 16777619;
            return (int)(hash & int.MaxValue);
        }

        private static bool Matches(byte[] expected, byte[] data, int start, int length) {
            if (expected.Length != length) return false;
            for (int index = 0; index < length; index++) {
                if (expected[index] != data[start + index]) return false;
            }
            return true;
        }
    }
}

/// <summary>A qualified name as an element or attribute is written: local name, prefix, folded namespace and exact URI.</summary>
internal readonly struct OfficeXmlName {
    internal OfficeXmlName(string localName, string? prefix, int namespaceId, int uriId) {
        LocalName = localName;
        Prefix = prefix;
        NamespaceId = namespaceId;
        UriId = uriId;
    }

    internal string LocalName { get; }

    /// <summary>The prefix as written, empty for none, or null for a created name whose prefix the writer chooses.</summary>
    internal string? Prefix { get; }

    internal int NamespaceId { get; }
    internal int UriId { get; }
}
