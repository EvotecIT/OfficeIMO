using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private void ValidatePublication(XDocument package, IReadOnlyDictionary<string, byte[]> entries, List<OfficeConversionFidelityDiagnostic> diagnostics, CancellationToken token) {
        XElement root = package.Root!;
        XElement metadata = root.Element(Opf + "metadata")!;
        if (!metadata.Elements(Dc + "title").Any(element => !string.IsNullOrWhiteSpace(element.Value)) ||
            !metadata.Elements(Dc + "language").Any(element => !string.IsNullOrWhiteSpace(element.Value)) ||
            !metadata.Elements(Dc + "identifier").Any(element => (string?)element.Attribute("id") == (string?)root.Attribute("unique-identifier") && !string.IsNullOrWhiteSpace(element.Value)))
            throw new InvalidDataException("Title, language and the selected package identifier are required.");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (XAttribute attribute in root.DescendantsAndSelf().Attributes("id")) {
            XmlConvert.VerifyNCName(attribute.Value);
            if (!ids.Add(attribute.Value)) throw new InvalidDataException("Duplicate package id: " + attribute.Value);
        }
        EpubManifestItem[] manifest = root.Element(Opf + "manifest")!.Elements(Opf + "item").Select(element => new EpubManifestItem(element, PackagePath)).ToArray();
        var byId = new Dictionary<string, EpubManifestItem>(StringComparer.Ordinal);
        foreach (EpubManifestItem item in manifest) {
            token.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(item.Id) || string.IsNullOrWhiteSpace(item.MediaType) || byId.ContainsKey(item.Id)) throw new InvalidDataException("Manifest ids and media types must be nonempty and unique.");
            byId.Add(item.Id, item);
            if ((HasToken(item.Properties, "scripted") || IsScriptMediaType(item.MediaType)) &&
                (!_originalEntries.TryGetValue(item.Reference.ContainerPath ?? string.Empty, out byte[]? retained) ||
                    !entries.TryGetValue(item.Reference.ContainerPath ?? string.Empty, out byte[]? current) || !retained.SequenceEqual(current)))
                throw new NotSupportedException("Script authoring is outside the EPUB writer contract.");
            EpubReference reference = item.Reference;
            if (reference.Kind == EpubReferenceKind.Container) {
                if (reference.ContainerPath == null || reference.Fragment != null || !entries.ContainsKey(reference.ContainerPath)) throw new InvalidDataException("Missing or fragmented manifest resource: " + item.Href);
            } else if (reference.Kind != EpubReferenceKind.External) throw new InvalidDataException("Invalid manifest reference: " + item.Href);
        }
        foreach (EpubManifestItem item in manifest) {
            foreach (string? target in new[] { item.FallbackId, item.MediaOverlayId }) {
                if (target != null && !byId.ContainsKey(target)) throw new InvalidDataException("Manifest relationship target missing: " + target);
            }
            var chain = new HashSet<string>(StringComparer.Ordinal) { item.Id };
            EpubManifestItem current = item;
            while (current.FallbackId != null) {
                current = byId[current.FallbackId];
                if (!chain.Add(current.Id)) throw new InvalidDataException("Cyclic manifest fallback chain.");
            }
        }
        XElement[] spine = root.Element(Opf + "spine")!.Elements(Opf + "itemref").ToArray();
        foreach (var repeated in spine.GroupBy(item => (string?)item.Attribute("idref"), StringComparer.Ordinal).Where(group => group.Count() > 1)) {
            if (_originalBytes == null) throw new InvalidDataException("Authored spine positions must have distinct manifest ids.");
            diagnostics.Add(new OfficeConversionFidelityDiagnostic("EPUB_WRITE_RETAINED_DUPLICATE_SPINE",
                "Repeated source reading positions were preserved; duplicate itemrefs do not conform to EPUB authoring requirements.",
                OfficeConversionLossKind.None, "OfficeIMO.Epub", PackagePath));
        }
        if (!spine.Any(item => (string?)item.Attribute("linear") != "no")) throw new InvalidDataException("At least one linear reading position is required.");
        foreach (XElement position in spine) {
            string id = (string?)position.Attribute("idref") ?? string.Empty;
            if (!byId.TryGetValue(id, out EpubManifestItem? item)) throw new InvalidDataException("Spine manifest id missing: " + id);
            while (item.MediaType != "application/xhtml+xml" && item.MediaType != "image/svg+xml" && item.FallbackId != null) item = byId[item.FallbackId];
            if (item.MediaType != "application/xhtml+xml" && item.MediaType != "image/svg+xml") throw new NotSupportedException("Spine item has no XHTML/SVG fallback: " + id);
        }
        string navPath = NavigationPath();
        XDocument navigation = ParseXml(entries[navPath], 64L * 1024 * 1024);
        if (PackageVersion == "3.0" && !navigation.Descendants(Html + "nav").Any(element => HasToken((string?)element.Attribute(Ops + "type"), "toc")))
            throw new InvalidDataException("An EPUB 3 TOC navigation section is required.");
        var anchors = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        var fragmented = new List<(string Owner, EpubReference Reference)>();
        foreach (EpubManifestItem item in manifest.Where(item => item.Reference.Kind == EpubReferenceKind.Container &&
            (item.MediaType == "application/xhtml+xml" || item.MediaType == "image/svg+xml" || item.MediaType == "application/x-dtbncx+xml"))) {
            token.ThrowIfCancellationRequested();
            string path = RequireLocalPath(item);
            if (_encryption.Any(encryption => encryption.Path == path && encryption.RequiresDecryption)) continue;
            XDocument content = path == navPath ? navigation : ParseXml(entries[path], 64L * 1024 * 1024);
            anchors[path] = new HashSet<string>(content.Descendants().Attributes().Where(attribute => attribute.Name == "id" ||
                attribute.Name == XNamespace.Xml + "id").Select(attribute => attribute.Value), StringComparer.Ordinal);
            bool rewritten = !_originalEntries.TryGetValue(path, out byte[]? original) || !original.SequenceEqual(entries[path]);
            if (rewritten && item.MediaType != "application/x-dtbncx+xml") {
                ValidateContent(content, item.MediaType);
                var resources = OfficeIMO.Html.HtmlResourcePipeline.BuildManifest(content.ToString(SaveOptions.DisableFormatting));
                if (resources.Resources.Any(resource => !resource.IsAllowed)) throw new NotSupportedException("Authored content contains a URL blocked by the shared HTML policy.");
                if (PackageVersion == "3.0") UpdateContentProperties(item, content, path, resources);
            }
            ValidateContentReferences(content, path, entries, fragmented, token);
        }
        foreach (XElement reference in root.Element(Opf + "guide")?.Elements(Opf + "reference") ?? Enumerable.Empty<XElement>())
            ValidateTarget(EpubReference.Resolve(PackagePath, (string?)reference.Attribute("href") ?? string.Empty), PackagePath, entries, fragmented);
        foreach (var target in fragmented) {
            token.ThrowIfCancellationRequested();
            if (anchors.TryGetValue(target.Reference.ContainerPath!, out HashSet<string>? idsInTarget) && !idsInTarget.Contains(target.Reference.Fragment!))
                throw new InvalidDataException("Content fragment missing in " + target.Owner + ": " + target.Reference.Original);
        }
    }

    private static void ValidateContentReferences(XDocument content, string owner, IReadOnlyDictionary<string, byte[]> entries,
        List<(string Owner, EpubReference Reference)> fragmented, CancellationToken token) {
        string? baseHref = content.Descendants(Html + "base").Select(element => (string?)element.Attribute("href")).FirstOrDefault(value => value != null);
        foreach (XElement element in content.Descendants()) {
            token.ThrowIfCancellationRequested();
            if (element.Name.Namespace != Html && element.Name.NamespaceName != "http://www.w3.org/2000/svg" && element.Name != Ncx + "content") continue;
            foreach (XAttribute attribute in element.Attributes().Where(attribute => attribute.Name == "href" || attribute.Name == "src" ||
                attribute.Name == "poster" || (element.Name == Html + "object" && attribute.Name == "data") ||
                attribute.Name == XName.Get("href", "http://www.w3.org/1999/xlink"))) {
                if (element.Name == Html + "base" || string.IsNullOrWhiteSpace(attribute.Value)) continue;
                EpubReference reference = EpubReference.Resolve(owner, baseHref, attribute.Value);
                ValidateTarget(reference, owner, entries, fragmented);
            }
            if (element.Name.Namespace == Html) foreach (var candidate in OfficeIMO.Html.HtmlSrcSetParser.Enumerate((string?)element.Attribute("srcset")))
                ValidateTarget(EpubReference.Resolve(owner, baseHref, candidate.Url), owner, entries, fragmented);
        }
    }

    private static void ValidateTarget(EpubReference reference, string owner, IReadOnlyDictionary<string, byte[]> entries,
        List<(string Owner, EpubReference Reference)> fragmented) {
        if (reference.Kind == EpubReferenceKind.Invalid) throw new InvalidDataException("Invalid content URL in " + owner);
        if (reference.Kind != EpubReferenceKind.Container) return;
        if (reference.ContainerPath == null || !entries.ContainsKey(reference.ContainerPath))
            throw new InvalidDataException("Content target missing in " + owner + ": " + reference.Original);
        if (!string.IsNullOrEmpty(reference.Fragment)) fragmented.Add((owner, reference));
    }

    private static void UpdateContentProperties(EpubManifestItem item, XDocument content, string owner, OfficeIMO.Html.HtmlResourceManifest resources) {
        // Linked CSS dependencies are not exhaustively traversed, so retain an explicit remote declaration.
        var properties = new List<string>(Tokens(item.Properties).Where(token => token != "svg" && token != "mathml"));
        if (content.Descendants().Any(element => element.Name.NamespaceName == "http://www.w3.org/2000/svg" && element.Name.LocalName == "svg") && item.MediaType == "application/xhtml+xml") properties.Add("svg");
        if (content.Descendants().Any(element => element.Name.NamespaceName == "http://www.w3.org/1998/Math/MathML")) properties.Add("mathml");
        // Shared HTML discovery also covers inline CSS, srcset, and non-hyperlink resource URLs.
        if (resources.Resources.Any(resource => resource.Kind != OfficeIMO.Html.HtmlResourceKind.Hyperlink &&
            EpubReference.Resolve(owner, resource.Source).Kind == EpubReferenceKind.External)) properties.Add("remote-resources");
        item.Properties = properties.Count == 0 ? null : string.Join(" ", properties.Distinct(StringComparer.Ordinal));
    }
}
