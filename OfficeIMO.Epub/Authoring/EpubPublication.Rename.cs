using OfficeIMO.Html;
using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private void VerifyAvailableContainerPath(string newPath, string? excludedPath = null) {
        string normalizedDestination = newPath.Normalize(NormalizationForm.FormC);
        if (_entries.Keys.Concat(new[] { PackagePath }).Where(path => path != excludedPath).Any(path => {
            string existing = path.Normalize(NormalizationForm.FormC);
            return string.Equals(existing.TrimEnd('/'), normalizedDestination, StringComparison.OrdinalIgnoreCase) ||
                existing.StartsWith(normalizedDestination + "/", StringComparison.OrdinalIgnoreCase) ||
                (!existing.EndsWith("/", StringComparison.Ordinal) && normalizedDestination.StartsWith(existing + "/", StringComparison.OrdinalIgnoreCase));
        }))
            throw new ArgumentException("The destination collides with an existing container path.", nameof(newPath));
    }

    private readonly Dictionary<string, string> _entryOrigins;
    private delegate (string Path, string? Fragment) ContentReferenceMap(EpubReference reference);
    /// <summary>
    /// Moves a local resource and repairs standard package, XHTML, SVG, NCX, SMIL and CSS references
    /// atomically. Manifest identifiers and spine positions remain unchanged. Requires one rendition,
    /// inspectable non-scripted resources and no encrypted resources or XML base declarations.
    /// </summary>
    public void RenameResource(string manifestId, string containerPath, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        string oldPath = RequireLocalPath(RequireManifestItem(manifestId));
        string newPath = VerifyContentPath(containerPath);
        EnsureResourceMutationAllowed(oldPath);
        EnsureResourceMutationAllowed(newPath);
        if (oldPath == newPath) return;
        if (_rootfilePaths.Length != 1 || _encryption.Count != 0)
            throw new NotSupportedException("Resource renaming requires one rendition with no encrypted or obfuscated resources.");
        VerifyAvailableContainerPath(newPath, oldPath);
        EpubManifestItem[] manifest = Manifest.ToArray();
        if (manifest.Any(item => HasToken(item.Properties, "scripted") || IsScriptMediaType(item.MediaType)))
            throw new NotSupportedException("Scripted publications cannot be renamed safely.");
        var entries = new Dictionary<string, byte[]>(_entries, StringComparer.Ordinal);
        var package = new XDocument(_package);
        var packageEdits = new List<(XAttribute Original, string Value)>();
        XAttribute[] originalAttributes = Root.DescendantsAndSelf().Attributes().ToArray();
        XAttribute[] proposedAttributes = package.Root!.DescendantsAndSelf().Attributes().ToArray();
        var packageReferences = new HashSet<XAttribute>(PackageResourceReferences(package.Root).Concat(
            package.Root.Element(Opf + "manifest")!.Elements(Opf + "item").Attributes("href")));
        if (package.Root.DescendantsAndSelf().Attributes(XNamespace.Xml + "base").Any())
            throw new NotSupportedException("Resource renaming does not support XML base declarations.");
        for (int index = 0; index < proposedAttributes.Length; index++) {
            if (!packageReferences.Contains(proposedAttributes[index])) continue;
            string value = RewriteMovedReference(PackagePath, null, PackagePath, null, proposedAttributes[index].Value, oldPath, newPath);
            if (value == proposedAttributes[index].Value) continue;
            proposedAttributes[index].Value = value;
            packageEdits.Add((originalAttributes[index], value));
        }
        foreach (var group in manifest.Where(item => item.Reference.Kind == EpubReferenceKind.Container).GroupBy(RequireLocalPath, StringComparer.Ordinal)) {
            cancellationToken.ThrowIfCancellationRequested();
            string path = group.Key, destination = path == oldPath ? newPath : path;
            string[] types = group.Select(item => item.MediaType).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            if (types.Length != 1) throw new NotSupportedException("Renaming requires an unambiguous media type for every resource.");
            if (!_entries.TryGetValue(path, out byte[]? payload)) throw new InvalidDataException("Resource is missing: " + path);
            byte[] rewritten = RewritePublicationResource(types[0], payload, path, destination, oldPath, newPath, cancellationToken);
            if (rewritten.LongLength > _maximumEntryBytes) throw new InvalidDataException("Renamed content exceeds its entry-byte limit.");
            if (!ReferenceEquals(payload, rewritten)) EnsureResourceMutationAllowed(path);
            entries[destination] = rewritten;
        }
        entries.Remove(oldPath);
        long delta = entries.Values.Sum(value => value.LongLength) - _retainedBytes;
        EnsurePackageBudget(package, delta);
        ValidatePublication(package, entries, new List<OfficeConversionFidelityDiagnostic>(), cancellationToken, changed: true);
        EnsurePackageBudget(package, delta);
        cancellationToken.ThrowIfCancellationRequested();
        if (_entryOrigins.TryGetValue(oldPath, out string? origin)) {
            _entryOrigins.Remove(oldPath);
            _entryOrigins.Add(newPath, origin);
        }
        foreach (var edit in packageEdits) edit.Original.Value = edit.Value;
        _entries.Clear();
        foreach (var entry in entries) _entries.Add(entry.Key, entry.Value);
        _retainedBytes += delta;
        MarkChanged();
    }

    private static bool IsRenameXml(string type) => HasMediaType(type, "application/xhtml+xml") || HasMediaType(type, "image/svg+xml") ||
        HasMediaType(type, "application/x-dtbncx+xml") || HasMediaType(type, "application/smil+xml");

    private static bool IsRenameLeaf(string type) => new[] { "image/png", "image/jpeg", "image/gif", "image/webp", "image/avif",
        "audio/mpeg", "audio/mp4", "audio/ogg", "audio/wav", "video/mp4", "video/webm", "video/ogg", "font/ttf", "font/otf", "font/woff", "font/woff2",
        "application/vnd.ms-opentype", "application/font-sfnt", "application/font-woff", "text/plain" }.Contains(type, StringComparer.OrdinalIgnoreCase);

    private static string RewriteMovedReference(string owner, string? oldBase, string destination, string? newBase, string value, string oldPath, string newPath, ContentReferenceMap? map = null, bool allowEmptyDocumentLink = false) {
        bool empty = string.IsNullOrWhiteSpace(value);
        if (empty && !allowEmptyDocumentLink) return value;
        // Empty hyperlink hrefs resolve to the current base document; the general resource
        // resolver intentionally rejects empty resource declarations. Keep that distinction.
        string referenceValue = empty ? "#" : value;
        EpubReference original = EpubReference.Resolve(owner, oldBase, referenceValue);
        if (!original.IsValid) throw new InvalidDataException("Invalid resource URL in " + owner + ": " + value);
        if (original.Kind != EpubReferenceKind.Container) return value;
        var mapped = map == null ? (Path: original.ContainerPath == oldPath ? newPath : original.ContainerPath!, Fragment: original.Fragment) : map(original);
        string target = mapped.Path;
        EpubReference current = EpubReference.Resolve(destination, newBase, referenceValue);
        if (current.Kind == EpubReferenceKind.Container && current.ContainerPath == target && current.Query == original.Query && current.Fragment == mapped.Fragment) return value;
        string relativeOwner = destination;
        if (!string.IsNullOrWhiteSpace(newBase)) {
            EpubReference directory = EpubReference.Resolve(destination, newBase, ".");
            if (directory.Kind != EpubReferenceKind.Container) throw new NotSupportedException("Cannot generate a container reference under an external base.");
            relativeOwner = directory.ContainerPath!.Length == 0 ? string.Empty : directory.ContainerPath + "/";
        }
        return RelativeHref(relativeOwner, target) + (original.Query == null ? string.Empty : "?" + original.Query) +
            (mapped.Fragment == null ? string.Empty : "#" + Uri.EscapeDataString(mapped.Fragment));
    }

    private static string RewriteMovedCss(string css, Func<string, string> rewrite, bool includeFragmentReferences = false) {
        bool changed = false;
        string result = HtmlResourcePipeline.RewriteCssResourceUrls(css, (value, _) => {
            string replacement = rewrite(value); changed |= replacement != value; return replacement;
        }, includeFragmentReferences);
        return changed ? result : css;
    }
}
