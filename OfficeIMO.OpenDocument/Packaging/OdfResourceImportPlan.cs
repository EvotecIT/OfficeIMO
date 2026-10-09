namespace OfficeIMO.OpenDocument;

/// <summary>Stages referenced binary package entries without replacing destination entries or reading external resources.</summary>
internal sealed class OdfResourceImportPlan {
    private readonly OdfPackage _source;
    private readonly OdfPackage _destination;
    private readonly HashSet<string> _paths;
    private readonly Dictionary<string, string> _mapped = new Dictionary<string, string>(StringComparer.Ordinal);
    private readonly List<(string Path, byte[] Bytes, string MediaType)> _entries = new List<(string, byte[], string)>();
    private int _next = 1;
    internal OdfResourceImportPlan(OdfPackage source, OdfPackage destination) {
        _source = source; _destination = destination;
        _paths = new HashSet<string>(destination.Entries.Select(entry => entry.Name), StringComparer.OrdinalIgnoreCase);
    }

    internal void Rewrite(XElement root) {
        foreach (XElement element in root.DescendantsAndSelf()) {
            XAttribute? href = element.Attribute(OdfNamespaces.XLink + "href");
            if (href == null || href.Value.StartsWith("#", StringComparison.Ordinal)) continue;
            bool resource = element.Name == OdfNamespaces.Draw + "image" || element.Name == OdfNamespaces.Draw + "fill-image" ||
                element.Name == OdfNamespaces.Style + "background-image" || element.Name == OdfNamespaces.Svg + "font-face-uri";
            if (Uri.TryCreate(href.Value, UriKind.Absolute, out Uri? absolute) && !absolute.IsFile) {
                if (resource) throw new NotSupportedException("Drawing import requires embedded resource bytes: " + href.Value);
                continue;
            }
            if (!resource) throw new NotSupportedException("Drawing import does not relocate relative external links or this resource kind: " + element.Name.LocalName);
            string sourcePath = OdfPackagePath.NormalizeHref(href.Value);
            if (!_mapped.TryGetValue(sourcePath, out string? destinationPath)) {
                OdfPackageEntry entry = _source.GetRequiredEntry(sourcePath);
                if (sourcePath.EndsWith("/", StringComparison.Ordinal) || string.IsNullOrEmpty(entry.MediaType))
                    throw new InvalidDataException("Imported resource is not a declared binary package entry: " + sourcePath);
                string extension = Path.GetExtension(sourcePath);
                if (extension.Any(character => !char.IsLetterOrDigit(character) && character != '.')) extension = ".bin";
                do { destinationPath = "Pictures/odgImport" + _next++.ToString(CultureInfo.InvariantCulture) + extension; }
                while (!_paths.Add(destinationPath));
                _entries.Add((destinationPath, entry.GetBytesForSave().ToArray(), entry.MediaType!));
                _mapped.Add(sourcePath, destinationPath);
            }
            int suffixAt = href.Value.IndexOfAny(new[] { '?', '#' });
            href.Value = destinationPath + (suffixAt < 0 ? "" : href.Value.Substring(suffixAt));
        }
    }

    internal void Apply() {
        foreach (var entry in _entries) _destination.AddOrReplaceEntry(entry.Path, entry.Bytes, entry.MediaType);
    }
}
