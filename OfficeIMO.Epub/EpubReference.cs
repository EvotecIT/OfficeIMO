namespace OfficeIMO.Epub;

/// <summary>
/// Represents an EPUB URL reference resolved against a container document path.
/// </summary>
public sealed class EpubReference {
    private readonly string? _encodedFragment;

    private EpubReference(
        string original,
        EpubReferenceKind kind,
        EpubReferenceError error,
        string? containerPath,
        string? externalUri,
        string? query,
        string? fragment,
        string? encodedFragment,
        bool isContainerRootRelative,
        bool isConforming) {
        Original = original;
        Kind = kind;
        Error = error;
        ContainerPath = containerPath;
        ContainerUrlPath = kind == EpubReferenceKind.Container && containerPath != null
            ? EncodeContainerPath(containerPath)
            : null;
        ExternalUri = externalUri;
        Query = query;
        Fragment = fragment;
        _encodedFragment = encodedFragment;
        IsContainerRootRelative = isContainerRootRelative;
        IsConforming = isConforming;
    }

    /// <summary>The original trimmed reference value.</summary>
    public string Original { get; }

    /// <summary>The resolved reference kind.</summary>
    public EpubReferenceKind Kind { get; }

    /// <summary>The resolution error, or <see cref="EpubReferenceError.None"/> on success.</summary>
    public EpubReferenceError Error { get; }

    /// <summary>Case-sensitive normalized archive path for a container reference.</summary>
    public string? ContainerPath { get; }

    /// <summary>
    /// URL-encoded form of <see cref="ContainerPath"/> for serialization in links and media references.
    /// </summary>
    public string? ContainerUrlPath { get; }

    /// <summary>External or data URL, including its query and fragment when present.</summary>
    public string? ExternalUri { get; }

    /// <summary>Raw query value without the leading question mark.</summary>
    public string? Query { get; }

    /// <summary>Decoded fragment value without the leading number sign.</summary>
    public string? Fragment { get; }

    /// <summary>Whether a container reference began at the virtual container root.</summary>
    public bool IsContainerRootRelative { get; }

    /// <summary>
    /// Whether the reference conforms to EPUB OCF URL restrictions. Safely resolved
    /// root-relative or backslash-separated container references are exposed with this
    /// value set to <see langword="false"/>.
    /// </summary>
    public bool IsConforming { get; }

    /// <summary>Whether the reference resolved successfully.</summary>
    public bool IsValid => Kind != EpubReferenceKind.Invalid;

    /// <summary>
    /// Resolved target without a fragment. Container targets include any original query.
    /// </summary>
    public string? Target {
        get {
            if (Kind == EpubReferenceKind.Container) {
                return AppendQuery(ContainerUrlPath, Query);
            }
            if (Kind == EpubReferenceKind.External || Kind == EpubReferenceKind.Data) {
                return RemoveFragment(ExternalUri);
            }
            return null;
        }
    }

    /// <summary>Fully resolved container path or preserved external/data URL.</summary>
    public string? ResolvedValue {
        get {
            if (Kind == EpubReferenceKind.Container) {
                return AppendFragment(AppendQuery(ContainerUrlPath, Query), _encodedFragment);
            }
            return ExternalUri;
        }
    }

    /// <summary>Resolves a reference against an EPUB container document path.</summary>
    /// <param name="basePath">Case-sensitive archive path of the containing document.</param>
    /// <param name="reference">URL reference to resolve.</param>
    /// <returns>A typed result. Invalid and unsafe values do not throw.</returns>
    public static EpubReference Resolve(string basePath, string reference) {
        string original = reference?.Trim() ?? string.Empty;
        if (original.Length == 0) return Invalid(original, EpubReferenceError.Empty);
        if (ContainsControlCharacter(original)) return Invalid(original, EpubReferenceError.ControlCharacter);
        if (!TryNormalizeBasePath(basePath, out string normalizedBase)) {
            return Invalid(original, EpubReferenceError.InvalidBasePath);
        }

        SplitReference(original, out string path, out string? query, out string? encodedFragment);
        string? fragment = DecodeFragment(encodedFragment);
        string pathAndQuery = AppendQuery(path, query) ?? string.Empty;
        string completeValue = AppendFragment(pathAndQuery, encodedFragment) ?? string.Empty;

        if (path.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) {
            return Valid(
                original,
                EpubReferenceKind.Data,
                null,
                completeValue,
                query,
                fragment,
                encodedFragment,
                false,
                true);
        }
        if (path.StartsWith("file:", StringComparison.OrdinalIgnoreCase) || LooksLikeWindowsDrivePath(path)) {
            return Invalid(original, EpubReferenceError.FileUrl);
        }
        if (path.StartsWith("//", StringComparison.Ordinal)) {
            return Valid(
                original,
                EpubReferenceKind.External,
                null,
                completeValue,
                query,
                fragment,
                encodedFragment,
                false,
                true);
        }
        if (!path.StartsWith("/", StringComparison.Ordinal) &&
            Uri.TryCreate(pathAndQuery, UriKind.Absolute, out Uri? absoluteUri)) {
            if (absoluteUri.IsFile) return Invalid(original, EpubReferenceError.FileUrl);
            return Valid(
                original,
                EpubReferenceKind.External,
                null,
                completeValue,
                query,
                fragment,
                encodedFragment,
                false,
                true);
        }

        bool hasRawBackslash = path.IndexOf('\\') >= 0;
        if (hasRawBackslash) path = path.Replace('\\', '/');
        bool isRootRelative = path.StartsWith("/", StringComparison.Ordinal);
        string relativePath = isRootRelative ? path.TrimStart('/') : path;
        var segments = new List<string>();
        if (!isRootRelative) {
            int lastSlash = normalizedBase.LastIndexOf('/');
            string baseDirectory = lastSlash < 0 ? string.Empty : normalizedBase.Substring(0, lastSlash);
            if (baseDirectory.Length > 0) segments.AddRange(baseDirectory.Split('/'));
        }

        if (relativePath.Length == 0) {
            if (isRootRelative) {
                return Valid(
                    original,
                    EpubReferenceKind.Container,
                    string.Empty,
                    null,
                    query,
                    fragment,
                    encodedFragment,
                    true,
                    false);
            }
            if (!TryNormalizeBasePath(normalizedBase, out string currentPath)) {
                return Invalid(original, EpubReferenceError.InvalidBasePath);
            }
            return Valid(
                original,
                EpubReferenceKind.Container,
                currentPath,
                null,
                query,
                fragment,
                encodedFragment,
                false,
                true);
        }

        string[] sourceSegments = relativePath.Split(new[] { '/' }, StringSplitOptions.None);
        foreach (string sourceSegment in sourceSegments) {
            if (sourceSegment.Length == 0 || IsSingleDotSegment(sourceSegment)) continue;
            if (IsDoubleDotSegment(sourceSegment)) {
                if (segments.Count == 0) return Invalid(original, EpubReferenceError.EscapesContainer);
                segments.RemoveAt(segments.Count - 1);
                continue;
            }
            if (!TryDecodePathSegment(sourceSegment, out string decodedSegment)) {
                return Invalid(original, EpubReferenceError.InvalidPath);
            }
            segments.Add(decodedSegment);
        }

        if (segments.Count == 0) {
            return Valid(
                original,
                EpubReferenceKind.Container,
                string.Empty,
                null,
                query,
                fragment,
                encodedFragment,
                isRootRelative,
                !isRootRelative && !hasRawBackslash);
        }
        return Valid(
            original,
            EpubReferenceKind.Container,
            string.Join("/", segments),
            null,
            query,
            fragment,
            encodedFragment,
            isRootRelative,
            !isRootRelative && !hasRawBackslash);
    }

    /// <summary>
    /// Resolves a content reference using an optional HTML <c>base</c> element value.
    /// </summary>
    /// <param name="documentPath">Case-sensitive archive path of the content document.</param>
    /// <param name="baseHref">Optional HTML base URL.</param>
    /// <param name="reference">URL reference to resolve.</param>
    /// <returns>A typed result. Invalid and unsafe values do not throw.</returns>
    public static EpubReference Resolve(string documentPath, string? baseHref, string reference) {
        if (string.IsNullOrWhiteSpace(baseHref)) return Resolve(documentPath, reference);

        string original = reference?.Trim() ?? string.Empty;
        if (original.Length == 0) return Invalid(original, EpubReferenceError.Empty);
        if (ContainsControlCharacter(original)) return Invalid(original, EpubReferenceError.ControlCharacter);

        EpubReference resolvedBase = Resolve(documentPath, baseHref!);
        if (!resolvedBase.IsValid) return Invalid(original, EpubReferenceError.InvalidBasePath);
        if (resolvedBase.Kind == EpubReferenceKind.Data) {
            return Invalid(original, EpubReferenceError.InvalidBasePath);
        }
        if (resolvedBase.Kind == EpubReferenceKind.External) {
            string externalBase = resolvedBase.ResolvedValue ?? string.Empty;
            if (externalBase.StartsWith("//", StringComparison.Ordinal)) externalBase = "https:" + externalBase;
            if (!Uri.TryCreate(externalBase, UriKind.Absolute, out Uri? baseUri) ||
                !Uri.TryCreate(baseUri, original, out Uri? resolvedExternal)) {
                return Invalid(original, EpubReferenceError.InvalidPath);
            }
            return Resolve(documentPath, resolvedExternal.AbsoluteUri);
        }

        string effectiveBase = resolvedBase.ContainerPath ?? documentPath;
        SplitReference(baseHref!.Trim(), out string basePathPart, out _, out _);
        SplitReference(original, out string referencePathPart, out string? query, out string? encodedFragment);
        if (referencePathPart.Length == 0) {
            string containerPath = basePathPart.EndsWith("/", StringComparison.Ordinal)
                ? effectiveBase.TrimEnd('/') + "/"
                : effectiveBase;
            return Valid(
                original,
                EpubReferenceKind.Container,
                containerPath,
                null,
                query ?? resolvedBase.Query,
                DecodeFragment(encodedFragment),
                encodedFragment,
                false,
                resolvedBase.IsConforming);
        }
        if (basePathPart.EndsWith("/", StringComparison.Ordinal)) {
            effectiveBase = effectiveBase.Length == 0
                ? "__officeimo_epub_base__"
                : effectiveBase.TrimEnd('/') + "/__officeimo_epub_base__";
        }
        EpubReference result = Resolve(effectiveBase, original);
        return resolvedBase.IsConforming || !result.IsValid
            ? result
            : result.WithConformance(false);
    }

    /// <inheritdoc />
    public override string ToString() => ResolvedValue ?? Original;

    private EpubReference WithConformance(bool isConforming) => new EpubReference(
        Original,
        Kind,
        Error,
        ContainerPath,
        ExternalUri,
        Query,
        Fragment,
        _encodedFragment,
        IsContainerRootRelative,
        isConforming);

    private static EpubReference Valid(
        string original,
        EpubReferenceKind kind,
        string? containerPath,
        string? externalUri,
        string? query,
        string? fragment,
        string? encodedFragment,
        bool isContainerRootRelative,
        bool isConforming) => new EpubReference(
            original,
            kind,
            EpubReferenceError.None,
            containerPath,
            externalUri,
            query,
            fragment,
            encodedFragment,
            isContainerRootRelative,
            isConforming);

    private static EpubReference Invalid(string original, EpubReferenceError error) => new EpubReference(
        original,
        EpubReferenceKind.Invalid,
        error,
        null,
        null,
        null,
        null,
        null,
        false,
        false);

    private static void SplitReference(string value, out string path, out string? query, out string? fragment) {
        int hash = value.IndexOf('#');
        string withoutFragment = hash < 0 ? value : value.Substring(0, hash);
        fragment = hash < 0 ? null : value.Substring(hash + 1);
        int question = withoutFragment.IndexOf('?');
        path = question < 0 ? withoutFragment : withoutFragment.Substring(0, question);
        query = question < 0 ? null : withoutFragment.Substring(question + 1);
    }

    private static string? AppendQuery(string? value, string? query) => value == null
        ? null
        : query == null ? value : value + "?" + query;

    private static string? AppendFragment(string? value, string? fragment) => value == null
        ? null
        : fragment == null ? value : value + "#" + fragment;

    private static string? RemoveFragment(string? value) {
        if (value == null) return null;
        int hash = value.IndexOf('#');
        return hash < 0 ? value : value.Substring(0, hash);
    }

    private static string? DecodeFragment(string? value) {
        if (value == null) return null;
        try {
            return Uri.UnescapeDataString(value);
        } catch {
            return value;
        }
    }

    private static string EncodeContainerPath(string value) => string.Join(
        "/",
        value.Split(new[] { '/' }, StringSplitOptions.None).Select(Uri.EscapeDataString));

    private static bool TryNormalizeBasePath(string? value, out string normalized) {
        normalized = value?.Replace('\\', '/').Trim() ?? string.Empty;
        if (normalized.Length == 0 || normalized.StartsWith("/", StringComparison.Ordinal)) return false;
        if (LooksLikeWindowsDrivePath(normalized) || Uri.TryCreate(normalized, UriKind.Absolute, out _)) return false;
        string[] segments = normalized.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        if (segments.Length == 0 || segments.Any(static segment => segment == "." || segment == ".." || ContainsControlCharacter(segment))) {
            return false;
        }
        normalized = string.Join("/", segments);
        return true;
    }

    private static bool TryDecodePathSegment(string value, out string decoded) {
        decoded = string.Empty;
        for (int index = 0; index < value.Length; index++) {
            if (value[index] != '%') continue;
            if (index + 2 >= value.Length || !Uri.IsHexDigit(value[index + 1]) || !Uri.IsHexDigit(value[index + 2])) {
                return false;
            }
            index += 2;
        }
        try {
            decoded = Uri.UnescapeDataString(value);
        } catch {
            return false;
        }
        if (decoded.Length == 0 || decoded.IndexOf('/') >= 0 || decoded.IndexOf('\\') >= 0 || ContainsControlCharacter(decoded)) {
            return false;
        }
        return true;
    }

    private static bool IsSingleDotSegment(string value) => string.Equals(value, ".", StringComparison.Ordinal)
        || string.Equals(value, "%2e", StringComparison.OrdinalIgnoreCase);

    private static bool IsDoubleDotSegment(string value) => string.Equals(value, "..", StringComparison.Ordinal)
        || string.Equals(value, ".%2e", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "%2e.", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "%2e%2e", StringComparison.OrdinalIgnoreCase);

    private static bool LooksLikeWindowsDrivePath(string value) => value.Length >= 2
        && char.IsLetter(value[0])
        && value[1] == ':';

    private static bool ContainsControlCharacter(string value) {
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (character <= '\u001F' || character == '\u007F') return true;
        }
        return false;
    }
}
