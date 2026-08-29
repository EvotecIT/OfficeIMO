namespace OfficeIMO.Email;

/// <summary>
/// Owns the immutable, ambiguity-aware identity policy used by HTML rewrites and resource resolution.
/// </summary>
internal sealed class EmailBodyResourceIdentityIndex {
    private const int NotFound = -1;
    private const int Ambiguous = -2;
    private readonly Entry[] _entries;

    internal EmailBodyResourceIdentityIndex(IReadOnlyList<EmailBodyResource> resources) {
        if (resources == null) throw new ArgumentNullException(nameof(resources));
        string?[] contentIds = resources
            .Select(resource => resource.ContentId)
            .ToArray();
        var contentIdCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var usedAliases = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string? contentId in contentIds) {
            if (string.IsNullOrWhiteSpace(contentId)) continue;
            usedAliases.Add(contentId!);
            contentIdCounts.TryGetValue(contentId!, out int count);
            contentIdCounts[contentId!] = count + 1;
        }

        _entries = new Entry[resources.Count];
        for (int index = 0; index < resources.Count; index++) {
            string? contentId = contentIds[index];
            string projectionContentId = !string.IsNullOrWhiteSpace(contentId) &&
                                         contentIdCounts[contentId!] == 1
                ? contentId!
                : CreateUniqueAlias(index, usedAliases);
            EmailBodyResource resource = resources[index];
            _entries[index] = new Entry(
                contentId,
                NormalizeReference(resource.ContentLocation),
                NormalizeReference(resource.FileName),
                projectionContentId);
        }
    }

    internal string? Rewrite(string value, Uri? baseUri) {
        if (value.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return value;
        int index;
        if (value.StartsWith("cid:", StringComparison.OrdinalIgnoreCase)) {
            string? contentId = DecodeContentId(value);
            index = contentId == null
                ? NotFound
                : FindUnique(entry => entry.MatchesContentId(contentId));
        } else {
            index = FindUnique(entry => MatchesLocation(value, null, entry.ContentLocation, baseUri));
            if (index == Ambiguous) return null;
            if (index == NotFound) {
                index = FindUnique(entry => MatchesFileName(
                    value, null, entry.FileName, baseUri));
            }
        }
        return index >= 0 ? "cid:" + _entries[index].ProjectionContentId : null;
    }

    internal EmailBodyResource? Resolve(
        IReadOnlyList<EmailBodyResource> resources,
        Uri? baseUri,
        string? reference,
        Uri? resolvedUri) {
        if (resources == null) throw new ArgumentNullException(nameof(resources));
        if (string.IsNullOrWhiteSpace(reference) && resolvedUri == null) return null;
        string value = (reference ?? resolvedUri!.OriginalString).Trim();
        int index;
        if (value.StartsWith("cid:", StringComparison.OrdinalIgnoreCase)) {
            string? contentId = DecodeContentId(value);
            index = contentId == null
                ? NotFound
                : FindUnique(entry => entry.MatchesContentId(contentId));
        } else {
            index = FindUnique(entry => MatchesLocation(value, resolvedUri, entry.ContentLocation, baseUri));
            if (index == Ambiguous) return null;
            if (index == NotFound) {
                index = FindUnique(entry => MatchesFileName(
                    value, resolvedUri, entry.FileName, baseUri));
            }
        }
        return index >= 0 && index < resources.Count ? resources[index] : null;
    }

    private int FindUnique(Func<Entry, bool> predicate) {
        int match = NotFound;
        for (int index = 0; index < _entries.Length; index++) {
            if (!predicate(_entries[index])) continue;
            if (match != NotFound) return Ambiguous;
            match = index;
        }
        return match;
    }

    private static bool MatchesLocation(
        string candidate,
        Uri? resolvedUri,
        string? location,
        Uri? baseUri) {
        if (location == null) return false;
        if (string.Equals(candidate, location, StringComparison.Ordinal)) return true;
        Uri? candidateUri = resolvedUri;
        if (candidateUri == null && !TryResolveUri(candidate, baseUri, out candidateUri)) return false;
        return TryResolveUri(location, baseUri, out Uri? locationUri) && candidateUri!.Equals(locationUri);
    }

    private static bool MatchesFileName(
        string candidate,
        Uri? resolvedUri,
        string? fileName,
        Uri? baseUri) {
        if (fileName == null) return false;
        if (string.Equals(candidate, fileName, StringComparison.OrdinalIgnoreCase)) return true;
        Uri? candidateUri = resolvedUri;
        if (candidateUri == null && !TryResolveUri(candidate, baseUri, out candidateUri)) return false;
        return TryResolveUri(fileName, baseUri, out Uri? fileNameUri) &&
            string.Equals(candidateUri!.AbsoluteUri, fileNameUri!.AbsoluteUri,
                StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryResolveUri(string value, Uri? baseUri, out Uri? uri) {
        if (baseUri != null && Uri.TryCreate(baseUri, value, out uri)) return true;
        return Uri.TryCreate(value, UriKind.Absolute, out uri);
    }

    private static string CreateUniqueAlias(int index, ISet<string> usedAliases) {
        string candidate = "officeimo-resource-" + index + "@officeimo.invalid";
        int suffix = 1;
        while (!usedAliases.Add(candidate)) {
            candidate = "officeimo-resource-" + index + "-" + suffix++ + "@officeimo.invalid";
        }
        return candidate;
    }

    private static string? DecodeContentId(string value) {
        try {
            return EmailBodyResource.NormalizeContentId(Uri.UnescapeDataString(value.Substring(4)));
        } catch (UriFormatException) {
            return null;
        }
    }

    private static string? NormalizeReference(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value!.Trim();

    private sealed class Entry {
        internal Entry(string? contentId, string? contentLocation, string? fileName,
            string projectionContentId) {
            ContentId = contentId;
            ContentLocation = contentLocation;
            FileName = fileName;
            ProjectionContentId = projectionContentId;
        }

        internal string? ContentId { get; }
        internal string? ContentLocation { get; }
        internal string? FileName { get; }
        internal string ProjectionContentId { get; }

        internal bool MatchesContentId(string value) =>
            string.Equals(ContentId, value, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(ProjectionContentId, value, StringComparison.OrdinalIgnoreCase);
    }
}
