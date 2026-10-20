using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Excel.Xlsb.Package;

internal sealed partial class XlsbPackagePartReader {
    private Dictionary<string, string>? _contentTypeOverrides;
    private Dictionary<string, string>? _contentTypeDefaults;

    internal string GetContentType(string partName) {
        if (_contentTypeOverrides == null) {
            byte[] bytes = ReadPart("[Content_Types].xml", _options.CancellationToken);
            using var input = new MemoryStream(bytes, writable: false);
            using var reader = XmlReader.Create(input, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = _options.MaxPartBytes
            });
            XNamespace ns = "http://schemas.openxmlformats.org/package/2006/content-types";
            XElement root = XDocument.Load(reader).Root ?? throw new InvalidDataException("Missing content-types root.");
            if (root.Name != ns + "Types") throw new InvalidDataException("Invalid package content-types root.");
            var overrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement item in root.Elements()) {
                _options.CancellationToken.ThrowIfCancellationRequested();
                if (item.Name != ns + "Override" && item.Name != ns + "Default") throw new InvalidDataException("Invalid content-types entry.");
                string? type = (string?)item.Attribute("ContentType");
                if (string.IsNullOrEmpty(type)) throw new InvalidDataException("Missing package part content type.");
                string? key = (string?)item.Attribute(item.Name == ns + "Override" ? "PartName" : "Extension");
                if (string.IsNullOrEmpty(key)) throw new InvalidDataException("Missing content-types entry key.");
                var entries = item.Name == ns + "Override" ? overrides : defaults;
                string entryKey = item.Name == ns + "Override" ? NormalizePartName(key!) : key!;
                if (entries.ContainsKey(entryKey)) throw new InvalidDataException("Duplicate content-types entry.");
                entries.Add(entryKey, type!);
            }
            _contentTypeOverrides = overrides; _contentTypeDefaults = defaults;
        }
        string normalized = NormalizePartName(partName);
        if (_contentTypeOverrides.TryGetValue(normalized, out string? explicitType)) return explicitType;
        if (_contentTypeDefaults!.TryGetValue(Path.GetExtension(normalized).TrimStart('.'), out string? defaultType)) return defaultType;
        throw new InvalidDataException($"The XLSB related part '{normalized}' has no declared content type.");
    }
}
