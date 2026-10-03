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
            EpubManifestItem? originalDeclaration = OriginalManifestItem(item.Id);
            if ((HasToken(item.Properties, "scripted") || IsScriptMediaType(item.MediaType) ||
                HasToken(originalDeclaration?.Properties, "scripted") || (originalDeclaration != null && IsScriptMediaType(originalDeclaration.MediaType))) &&
                !IsRetainedScriptDeclaration(item, entries))
                throw new NotSupportedException("Script authoring is outside the EPUB writer contract.");
            EpubReference reference = item.Reference;
            if (reference.Kind == EpubReferenceKind.Container) {
                if (reference.ContainerPath == null || reference.Fragment != null || !entries.ContainsKey(reference.ContainerPath)) throw new InvalidDataException("Missing or fragmented manifest resource: " + item.Href);
            } else if (reference.Kind != EpubReferenceKind.External) throw new InvalidDataException("Invalid manifest reference: " + item.Href);
        }
        foreach (EpubManifestItem item in manifest) {
            foreach (string? target in new[] { item.FallbackId, item.FallbackStyleId, item.MediaOverlayId }) {
                if (target != null && !byId.ContainsKey(target)) throw new InvalidDataException("Manifest relationship target missing: " + target);
            }
            var chain = new HashSet<string>(StringComparer.Ordinal) { item.Id };
            EpubManifestItem current = item;
            while (current.FallbackId != null) {
                current = byId[current.FallbackId];
                if (!chain.Add(current.Id)) throw new InvalidDataException("Cyclic manifest fallback chain.");
            }
        }
        foreach (XAttribute refinement in root.Descendants().Where(element => element.Name == Opf + "meta" || element.Name == Opf + "link").Attributes("refines")) {
            EpubReference reference = EpubReference.Resolve(PackagePath, refinement.Value);
            if (reference.Kind != EpubReferenceKind.Container || reference.ContainerPath == null)
                throw new InvalidDataException("Invalid metadata refinement: " + refinement.Value);
            if (reference.ContainerPath == PackagePath && (reference.Fragment == null || !ids.Contains(reference.Fragment)))
                throw new InvalidDataException("Metadata refinement target missing: " + refinement.Value);
            if (reference.ContainerPath != PackagePath && !entries.ContainsKey(reference.ContainerPath))
                throw new InvalidDataException("Metadata refinement resource missing: " + refinement.Value);
        }
        foreach (XElement binding in root.Element(Opf + "bindings")?.Elements(Opf + "mediaType") ?? Enumerable.Empty<XElement>())
            if (!byId.ContainsKey((string?)binding.Attribute("handler") ?? string.Empty)) throw new InvalidDataException("Binding handler manifest id missing.");
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
            while (!HasMediaType(item.MediaType, "application/xhtml+xml") && !HasMediaType(item.MediaType, "image/svg+xml") && item.FallbackId != null) item = byId[item.FallbackId];
            if (!HasMediaType(item.MediaType, "application/xhtml+xml") && !HasMediaType(item.MediaType, "image/svg+xml")) throw new NotSupportedException("Spine item has no XHTML/SVG fallback: " + id);
        }
        string navPath = NavigationPath();
        XDocument navigation = ParseXml(entries[navPath], _maximumEntryBytes);
        ValidateNavigationRoot(navigation);
        ValidateNavigationDocument(navigation, navPath, manifest, spine.Select(item => (string?)item.Attribute("idref") ?? string.Empty));
        var anchors = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        var fragmented = new List<(string Owner, EpubReference Reference)>();
        foreach (EpubManifestItem item in manifest.Where(item => item.Reference.Kind == EpubReferenceKind.Container &&
            (HasMediaType(item.MediaType, "application/xhtml+xml") || HasMediaType(item.MediaType, "image/svg+xml") || HasMediaType(item.MediaType, "application/x-dtbncx+xml")))) {
            token.ThrowIfCancellationRequested();
            string path = RequireLocalPath(item);
            if (_encryption.Any(encryption => encryption.Path == path && encryption.RequiresDecryption)) continue;
            XDocument content = path == navPath ? navigation : ParseXml(entries[path], _maximumEntryBytes);
            anchors[path] = new HashSet<string>(content.Descendants().Attributes().Where(attribute => attribute.Name == "id" ||
                attribute.Name == XNamespace.Xml + "id").Select(attribute => attribute.Value), StringComparer.Ordinal);
            bool rewritten = !_originalEntries.TryGetValue(path, out byte[]? original) || !original.SequenceEqual(entries[path]) ||
                !HasMediaType(OriginalManifestItem(item.Id)?.MediaType, item.MediaType);
            if (rewritten && !HasMediaType(item.MediaType, "application/x-dtbncx+xml")) {
                ValidateContent(content, item.MediaType);
                string analysisContent = content.ToString(SaveOptions.DisableFormatting);
                var limits = OfficeIMO.Html.HtmlConversionLimits.CreateUntrustedProfile();
                // ParseXml already bounded the retained bytes; canonical escaping can enlarge this snapshot.
                limits.MaxInputCharacters = (int)Math.Min(Math.Max(_maximumEntryBytes, analysisContent.Length), int.MaxValue);
                var resources = OfficeIMO.Html.HtmlResourcePipeline.BuildManifest(analysisContent,
                    new OfficeIMO.Html.HtmlResourcePipelineOptions { Limits = limits });
                if (resources.Resources.Any(resource => !resource.IsAllowed)) throw new NotSupportedException("Authored content contains a URL blocked by the shared HTML policy.");
                if (PackageVersion == "3.0") UpdateContentProperties(item, content, path, resources);
            }
            ValidateContentReferences(content, path, entries, fragmented, token);
        }
        foreach (XAttribute reference in PackageResourceReferences(root)) {
            EpubReference target = EpubReference.Resolve(PackagePath, reference.Value);
            if (target.ContainerPath != PackagePath) ValidateTarget(target, PackagePath, entries, fragmented);
            if (reference.Parent?.Name == Opf + "reference" || reference.Parent?.Name == Opf + "site")
                RequireSpineTarget(target, manifest, spine.Select(item => (string?)item.Attribute("idref") ?? string.Empty), requireContainer: true);
        }
        foreach (var target in fragmented) {
            token.ThrowIfCancellationRequested();
            if (anchors.TryGetValue(target.Reference.ContainerPath!, out HashSet<string>? idsInTarget) && !idsInTarget.Contains(target.Reference.Fragment!))
                throw new InvalidDataException("Content fragment missing in " + target.Owner + ": " + target.Reference.Original);
        }
    }

    private static void ValidateContentReferences(XDocument content, string owner, IReadOnlyDictionary<string, byte[]> entries,
        List<(string Owner, EpubReference Reference)> fragmented, CancellationToken token) {
        foreach (EpubReference reference in ContentResourceReferences(content, owner, token)) {
            token.ThrowIfCancellationRequested();
            ValidateTarget(reference, owner, entries, fragmented);
        }
    }

    private static IEnumerable<EpubReference> ContentResourceReferences(XDocument content, string owner, CancellationToken token = default) {
        string? baseHref = content.Root?.Name == Html + "html" ? content.Root.Element(Html + "head")?.Elements(Html + "base")
            .Select(element => (string?)element.Attribute("href")).FirstOrDefault(value => value != null) : null;
        foreach (XElement element in content.Descendants()) {
            token.ThrowIfCancellationRequested();
            if (element.Name.Namespace != Html && element.Name.NamespaceName != "http://www.w3.org/2000/svg" && element.Name != Ncx + "content") continue;
            foreach (XAttribute attribute in element.Attributes().Where(attribute => attribute.Name == "href" || attribute.Name == "src" ||
                attribute.Name == "poster" || (element.Name == Html + "object" && attribute.Name == "data") ||
                attribute.Name == XName.Get("href", "http://www.w3.org/1999/xlink"))) {
                if (element.Name == Html + "base" || string.IsNullOrWhiteSpace(attribute.Value)) continue;
                yield return EpubReference.Resolve(owner, baseHref, attribute.Value);
            }
            if (element.Name.Namespace == Html) foreach (var candidate in OfficeIMO.Html.HtmlSrcSetParser.Enumerate((string?)element.Attribute("srcset")))
                yield return EpubReference.Resolve(owner, baseHref, candidate.Url);
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
        if (content.Descendants().Any(element => element.Name.NamespaceName == "http://www.w3.org/2000/svg" && element.Name.LocalName == "svg") && HasMediaType(item.MediaType, "application/xhtml+xml")) properties.Add("svg");
        if (content.Descendants().Any(element => element.Name.NamespaceName == "http://www.w3.org/1998/Math/MathML")) properties.Add("mathml");
        // Shared HTML discovery also covers inline CSS, srcset, and non-hyperlink resource URLs.
        if (resources.Resources.Any(resource => resource.Kind != OfficeIMO.Html.HtmlResourceKind.Hyperlink &&
            EpubReference.Resolve(owner, resource.ResolvedSource.Length == 0 ? resource.Source : resource.ResolvedSource).Kind == EpubReferenceKind.External)) properties.Add("remote-resources");
        item.Properties = properties.Count == 0 ? null : string.Join(" ", properties.Distinct(StringComparer.Ordinal));
    }

    private EpubManifestItem? OriginalManifestItem(string id) => _originalManifest.TryGetValue(id, out EpubManifestItem? item) ? item : null;

    private bool IsRetainedScriptDeclaration(EpubManifestItem item, IReadOnlyDictionary<string, byte[]> entries) {
        EpubManifestItem? original = OriginalManifestItem(item.Id);
        if (_originalBytes == null || original == null || original.Href != item.Href || !HasMediaType(original.MediaType, item.MediaType) ||
            HasToken(original.Properties, "scripted") != HasToken(item.Properties, "scripted")) return false;
        if (item.Reference.Kind == EpubReferenceKind.External) return true;
        string path = item.Reference.ContainerPath ?? string.Empty;
        return _originalEntries.TryGetValue(path, out byte[]? retained) && entries.TryGetValue(path, out byte[]? current) && retained.SequenceEqual(current);
    }

    private bool ReferencesPackageId(string value, string id) {
        EpubReference reference = EpubReference.Resolve(PackagePath, value);
        return reference.Kind == EpubReferenceKind.Container && reference.ContainerPath == PackagePath && reference.Fragment == id;
    }

    private static IEnumerable<XAttribute> PackageResourceReferences(XElement root) => root.Descendants()
        .Where(element => element.Name == Opf + "link" || element.Name == Opf + "reference" || element.Name == Opf + "site")
        .Attributes("href").Concat(root.Descendants().Where(element => element.Name == Opf + "meta" || element.Name == Opf + "link").Attributes("refines"));

    private void ValidateNavigationDocument(XDocument navigation, string path, EpubManifestItem[] manifest, IEnumerable<string> spineIds) {
        if (PackageVersion == "2.0") {
            if (navigation.Root?.Element(Ncx + "navMap")?.Elements(Ncx + "navPoint").Any() != true)
                throw new InvalidDataException("A nonempty NCX table of contents is required.");
            foreach (XAttribute source in navigation.Descendants(Ncx + "content").Attributes("src"))
                RequireSpineTarget(EpubReference.Resolve(path, source.Value), manifest, spineIds, requireContainer: true);
            return;
        }
        XElement[] tocs = navigation.Descendants(Html + "nav").Where(element => HasToken((string?)element.Attribute(Ops + "type"), "toc")).ToArray();
        if (tocs.Length != 1 || tocs[0].Element(Html + "ol")?.Elements(Html + "li").Any() != true)
            throw new InvalidDataException("One nonempty EPUB 3 TOC navigation section is required.");
        foreach (XElement nav in navigation.Descendants(Html + "nav").Where(element => HasToken((string?)element.Attribute(Ops + "type"), "landmarks")))
            if (nav.Descendants(Html + "a").Any(anchor => string.IsNullOrWhiteSpace((string?)anchor.Attribute(Ops + "type"))))
                throw new InvalidDataException("Every landmark link requires a semantic type.");
        string? baseHref = navigation.Root?.Element(Html + "head")?.Elements(Html + "base").Select(element => (string?)element.Attribute("href")).FirstOrDefault();
        foreach (XElement anchor in navigation.Descendants().Where(element => element.Name == Html + "a" || element.Name == Html + "area")) {
            string? href = (string?)anchor.Attribute("href");
            bool ownedNavigation = anchor.Ancestors(Html + "nav").Any(nav => new[] { "toc", "page-list", "landmarks" }
                .Any(type => HasToken((string?)nav.Attribute(Ops + "type"), type)));
            if (href != null) RequireSpineTarget(EpubReference.Resolve(path, baseHref, href), manifest, spineIds, ownedNavigation);
        }
    }
}
