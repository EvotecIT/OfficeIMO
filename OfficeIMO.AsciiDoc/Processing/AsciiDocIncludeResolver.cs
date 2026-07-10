namespace OfficeIMO.AsciiDoc;

/// <summary>Request passed to an explicitly configured include resolver.</summary>
public sealed class AsciiDocIncludeRequest {
    internal AsciiDocIncludeRequest(string target, string? currentSourceName, int depth, AsciiDocDocumentAttributes attributes) {
        Target = target;
        CurrentSourceName = currentSourceName;
        Depth = depth;
        Attributes = attributes;
    }

    /// <summary>Attribute-expanded include target.</summary>
    public string Target { get; }

    /// <summary>Source containing the directive, when known.</summary>
    public string? CurrentSourceName { get; }

    /// <summary>One-based include nesting depth.</summary>
    public int Depth { get; }

    /// <summary>Attributes set when the include was encountered.</summary>
    public AsciiDocDocumentAttributes Attributes { get; }
}

/// <summary>Text returned by an include resolver.</summary>
public sealed class AsciiDocIncludeResult {
    /// <summary>Creates resolved include content.</summary>
    public AsciiDocIncludeResult(string content, string? sourceName = null) {
        Content = content ?? throw new ArgumentNullException(nameof(content));
        SourceName = sourceName;
    }

    /// <summary>Resolved characters.</summary>
    public string Content { get; }

    /// <summary>Stable source identifier used for nested resolution and cycle detection.</summary>
    public string? SourceName { get; }
}

/// <summary>Caller-controlled boundary for include resolution.</summary>
public interface IAsciiDocIncludeResolver {
    /// <summary>Resolves a target or returns null when unavailable or denied.</summary>
    AsciiDocIncludeResult? Resolve(AsciiDocIncludeRequest request);
}

/// <summary>BCL-only resolver confined to a configured filesystem root.</summary>
public sealed class AsciiDocRootedFileIncludeResolver : IAsciiDocIncludeResolver {
    private readonly string _root;
    private readonly string _rootPrefix;
    private readonly StringComparison _pathComparison;

    /// <summary>Creates a resolver. URI, absolute, escaping, and symbolic-link targets are denied.</summary>
    public AsciiDocRootedFileIncludeResolver(string rootDirectory) {
        if (rootDirectory == null) throw new ArgumentNullException(nameof(rootDirectory));
        string fullRoot = Path.GetFullPath(rootDirectory);
        string pathRoot = Path.GetPathRoot(fullRoot) ?? string.Empty;
        _root = fullRoot.Length > pathRoot.Length
            ? fullRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            : fullRoot;
        _rootPrefix = _root.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ||
                      _root.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)
            ? _root
            : _root + Path.DirectorySeparatorChar;
        _pathComparison = Path.DirectorySeparatorChar == '\\' ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
    }

    /// <summary>
    /// Allows filesystem reparse points and symbolic links. Disabled by default because they can escape lexical root checks.
    /// </summary>
    public bool AllowSymbolicLinks { get; set; }

    /// <inheritdoc />
    public AsciiDocIncludeResult? Resolve(AsciiDocIncludeRequest request) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (request.Target.Length == 0 || request.Target.IndexOf('\0') >= 0 || Path.IsPathRooted(request.Target)) return null;
        if (Uri.TryCreate(request.Target, UriKind.Absolute, out Uri? uri) && uri.IsAbsoluteUri) return null;

        string baseDirectory = GetBaseDirectory(request.CurrentSourceName);
        string candidate = Path.GetFullPath(Path.Combine(baseDirectory, request.Target));
        if (!IsWithinRoot(candidate) || !File.Exists(candidate)) return null;
        if (!AllowSymbolicLinks && ContainsReparsePoint(candidate)) return null;
        return new AsciiDocIncludeResult(File.ReadAllText(candidate), candidate);
    }

    private string GetBaseDirectory(string? sourceName) {
        if (sourceName == null) return _root;
        string fullSource = Path.GetFullPath(sourceName);
        if (!IsWithinRoot(fullSource)) return _root;
        return Path.GetDirectoryName(fullSource) ?? _root;
    }

    private bool IsWithinRoot(string path) =>
        string.Equals(path, _root, _pathComparison) ||
        path.StartsWith(_rootPrefix, _pathComparison);

    private bool ContainsReparsePoint(string path) {
        string relative = path.Substring(_root.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string current = _root;
        string[] parts = relative.Split(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);
        for (int index = 0; index < parts.Length; index++) {
            current = Path.Combine(current, parts[index]);
            if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0) return true;
        }
        return false;
    }
}
