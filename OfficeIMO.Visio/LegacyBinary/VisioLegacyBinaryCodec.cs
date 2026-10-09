using System.Globalization;
using System.Threading;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

// The binary adapter projects cached ShapeSheet values into the existing XML/model loader.
// Native formulas and unsupported carriers never become executable model behavior.
internal sealed partial class VisioLegacyBinaryCodec {
    private static readonly XNamespace Ns = VisioLegacyXmlCodec.Legacy;
    private readonly VisioBinaryContainer _container;
    private readonly VisioLegacyBinaryImportOptions _options;
    private readonly CancellationToken _token;
    private readonly Dictionary<uint, int> _unmapped = new();
    private readonly List<OfficeCompatibilityFinding> _findings = new();
    private int _items, _characters;

    internal VisioLegacyBinaryCodec(byte[] source, VisioLegacyBinaryImportOptions options, CancellationToken token) {
        _options = options; _token = token; _container = new(source, options, token);
    }

    internal XDocument Read() {
        var root = _container.ReadRoot();
        if (root.Type != 0x14) throw new InvalidDataException("Binary Visio has no document trailer pointer.");
        var document = new XElement(Ns + "VisioDocument", new XAttribute("version", "11.0"));
        foreach (var node in root.Children) {
            _token.ThrowIfCancellationRequested();
            if (node.Type == 0x16) document.Add(ReadColors(node));
            if (node.Type == 0xd8) document.Add(ReadFonts(node));
            if (node.Type == 0x1a) document.Add(ReadStyles(node));
        }
        var masters = new XElement(Ns + "Masters");
        var pages = new XElement(Ns + "Pages");
        foreach (var node in root.Children) {
            if (node.Type == 0x1d) {
                foreach (var master in node.Children.Where(child => child.Type == 0x1e))
                    masters.Add(ReadMaster(master, "Master-" + master.Id));
            } else if (node.Type == 0x27) {
                foreach (var page in node.Children.Where(child => child.Type == 0x15))
                    pages.Add(ReadPage(page, "Page-" + (page.Id + 1)));
            }
        }
        if (_items == 0) throw new InvalidDataException("Binary Visio has no shapes in the supported cached profile.");
        document.Add(masters, pages);
        return new XDocument(document);
    }

    internal OfficeLegacyImportReport CreateReport(VisioXmlConversionReport normalization) {
        _findings.Add(new OfficeCompatibilityFinding("VSD_CACHED_RECONSTRUCTION", "ShapeSheet",
            "Version 11 cached transforms, supported geometry, styles and text are reconstructed. Native formulas, recalculation, page/master/shape names, advanced styling, custom data and binary carriers are not preserved; this is an editable conversion rather than binary round-trip support.",
            OfficeCompatibilityState.Approximated, OfficeCompatibilitySeverity.Warning,
            OfficeCompatibilityImpact.Editability | OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Visual, representsLoss: true));
        foreach (var entry in _unmapped.OrderBy(pair => pair.Key))
            _findings.Add(new OfficeCompatibilityFinding("VSD_RECORD_OMITTED", "BinaryRecord",
                $"Record type 0x{entry.Key:X} is outside the cached import profile ({entry.Value} records).",
                OfficeCompatibilityState.Dropped, OfficeCompatibilitySeverity.Warning,
                OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Visual, representsLoss: true,
                sourceLocation: "record:0x" + entry.Key.ToString("X", CultureInfo.InvariantCulture)));
        foreach (var diagnostic in normalization.FidelityDiagnostics)
            _findings.Add(new OfficeCompatibilityFinding(diagnostic.Code, "ModelProjection", diagnostic.Message,
                diagnostic.LossKind == OfficeConversionLossKind.Omission ? OfficeCompatibilityState.Dropped : OfficeCompatibilityState.Approximated,
                OfficeCompatibilitySeverity.Warning, OfficeCompatibilityImpact.Visual | OfficeCompatibilityImpact.Editability,
                representsLoss: diagnostic.LossKind != OfficeConversionLossKind.None, sourceLocation: diagnostic.Location));
        OfficeLegacyInertContentKind inert = OfficeLegacyInertContentKind.None;
        foreach (var entry in _container.DirectoryEntries) {
            string name = entry.Name;
            if (name.IndexOf("VBA", StringComparison.OrdinalIgnoreCase) >= 0 || name.Equals("_VBA_PROJECT", StringComparison.OrdinalIgnoreCase))
                inert |= OfficeLegacyInertContentKind.Macros;
            if (name.IndexOf("ObjectPool", StringComparison.OrdinalIgnoreCase) >= 0)
                inert |= OfficeLegacyInertContentKind.EmbeddedObjects;
        }
        if (_unmapped.ContainsKey(0x1f) || _unmapped.ContainsKey(0x0c)) inert |= OfficeLegacyInertContentKind.EmbeddedObjects;
        if (_unmapped.ContainsKey(0xc4)) inert |= OfficeLegacyInertContentKind.ExternalLinks;
        return new OfficeLegacyImportReport("visio-binary-v11", OfficeLegacyImportQuality.Structured, _findings, inert, _items);
    }

    private List<VisioBinaryChunks.Chunk> Chunks(VisioBinaryContainer.Node node) {
        if ((node.Format >> 4) is not (8 or 12 or 13)) throw new InvalidDataException("Binary Visio shape content is not a record stream.");
        return VisioBinaryChunks.Read(node, _container, _options.MaxDepth).ToList();
    }
    private void Item() {
        _token.ThrowIfCancellationRequested();
        if (++_items > _options.Limits.MaxItems) throw new InvalidDataException("Binary Visio projected-item budget exceeded.");
    }
    private string Text(string value) {
        if (value.Length > _options.Limits.MaxTextCharacters - _characters) throw new InvalidDataException("Binary Visio text budget exceeded.");
        _characters += value.Length;
        return value;
    }
    private void Unmapped(uint type) { _unmapped.TryGetValue(type, out int count); _unmapped[type] = count + 1; }
    private static XElement Cell(string name, double value) => new(Ns + name, value.ToString("R", CultureInfo.InvariantCulture));
    private static XElement Cell(string name, string value) => new(Ns + name, value);
    private static XAttribute Id(uint value) => new("ID", value.ToString(CultureInfo.InvariantCulture));
}
