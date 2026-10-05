using OfficeIMO.Html;
using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private readonly Dictionary<string, string> _entryOrigins;
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
        string normalizedDestination = newPath.Normalize(NormalizationForm.FormC);
        if (_entries.Keys.Concat(new[] { PackagePath }).Where(path => path != oldPath).Any(path => {
            string existing = path.Normalize(NormalizationForm.FormC);
            return string.Equals(existing.TrimEnd('/'), normalizedDestination, StringComparison.OrdinalIgnoreCase) ||
                existing.StartsWith(normalizedDestination + "/", StringComparison.OrdinalIgnoreCase) ||
                (!existing.EndsWith("/", StringComparison.Ordinal) && normalizedDestination.StartsWith(existing + "/", StringComparison.OrdinalIgnoreCase));
        }))
            throw new ArgumentException("The destination collides with an existing container path.", nameof(containerPath));
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
            string type = types[0];
            byte[] rewritten = payload;
            if (HasMediaType(type, "text/css")) {
                if (!HtmlResourcePipeline.TryDecodeStylesheet(payload, "text/css", out string css)) throw new InvalidDataException("Stylesheet cannot be decoded: " + path);
                string result = RewriteMovedCss(css, value => RewriteMovedReference(path, null, destination, null, value, oldPath, newPath));
                if (result != css) {
                    // The shared decoder has consumed any source encoding. Emit matching UTF-8 bytes and declaration.
                    if (result.TrimStart().StartsWith("@charset", StringComparison.OrdinalIgnoreCase)) {
                        int end = result.IndexOf(';');
                        if (end >= 0) result = "@charset \"UTF-8\";" + result.Substring(end + 1);
                    }
                    rewritten = new UTF8Encoding(false, true).GetBytes(result);
                }
            } else if (IsRenameXml(type)) {
                XDocument document = ParseXml(payload, _maximumEntryBytes);
                XName expected = HasMediaType(type, "application/xhtml+xml") ? Html + "html" : HasMediaType(type, "image/svg+xml") ?
                    XName.Get("svg", "http://www.w3.org/2000/svg") : HasMediaType(type, "application/smil+xml") ?
                    XName.Get("smil", "http://www.w3.org/ns/SMIL") : Ncx + "ncx";
                if (document.Root?.Name != expected) throw new InvalidDataException("Resource XML root does not match its media type: " + path);
                if (RewriteMovedXml(document, path, destination, oldPath, newPath, cancellationToken))
                    rewritten = SerializeXml(document, _maximumEntryBytes);
            } else if (!IsRenameLeaf(type)) {
                throw new NotSupportedException("Resource renaming cannot inspect references in media type: " + type);
            }
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

    private static string RewriteMovedReference(string owner, string? oldBase, string destination, string? newBase, string value, string oldPath, string newPath) {
        if (string.IsNullOrWhiteSpace(value)) return value;
        EpubReference original = EpubReference.Resolve(owner, oldBase, value);
        if (!original.IsValid) throw new InvalidDataException("Invalid resource URL in " + owner + ": " + value);
        if (original.Kind != EpubReferenceKind.Container) return value;
        string target = original.ContainerPath == oldPath ? newPath : original.ContainerPath!;
        EpubReference current = EpubReference.Resolve(destination, newBase, value);
        if (current.Kind == EpubReferenceKind.Container && current.ContainerPath == target && current.Query == original.Query && current.Fragment == original.Fragment) return value;
        string relativeOwner = destination;
        if (!string.IsNullOrWhiteSpace(newBase)) {
            EpubReference directory = EpubReference.Resolve(destination, newBase, ".");
            if (directory.Kind != EpubReferenceKind.Container) throw new NotSupportedException("Cannot generate a container reference under an external base.");
            relativeOwner = directory.ContainerPath!.Length == 0 ? string.Empty : directory.ContainerPath + "/";
        }
        return RelativeHref(relativeOwner, target) + (original.Query == null ? string.Empty : "?" + original.Query) +
            (original.Fragment == null ? string.Empty : "#" + Uri.EscapeDataString(original.Fragment));
    }

    private static string RewriteMovedCss(string css, Func<string, string> rewrite) {
        bool changed = false;
        string result = HtmlResourcePipeline.RewriteCssResourceUrls(css, (value, _) => {
            string replacement = rewrite(value); changed |= replacement != value; return replacement;
        });
        return changed ? result : css;
    }
}
