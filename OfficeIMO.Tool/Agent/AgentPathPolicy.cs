namespace OfficeIMO.Tool.Agent;

internal sealed class AgentPathPolicy {
    internal const string AllowedRootsEnvironmentVariable = "OFFICEIMO_MCP_ALLOWED_ROOTS";
    private readonly IReadOnlyList<string> _allowedRoots;

    internal AgentPathPolicy(IEnumerable<string>? allowedRoots = null) {
        _allowedRoots = (allowedRoots ?? Array.Empty<string>())
            .Where(static root => !string.IsNullOrWhiteSpace(root))
            .Select(OfficeImoToolPathSafety.ResolveExistingLinks)
            .Distinct(StringComparer.Ordinal)
            .ToArray();
    }

    internal static AgentPathPolicy FromEnvironment() {
        string? configured = Environment.GetEnvironmentVariable(AllowedRootsEnvironmentVariable);
        return string.IsNullOrWhiteSpace(configured)
            ? new AgentPathPolicy()
            : FromConfiguredRoots(configured);
    }

    internal static AgentPathPolicy FromMcpEnvironment() =>
        ForMcp(
            Environment.GetEnvironmentVariable(AllowedRootsEnvironmentVariable),
            Directory.GetCurrentDirectory());

    internal static AgentPathPolicy ForMcp(string? configuredRoots, string workingDirectory) {
        if (!string.IsNullOrWhiteSpace(configuredRoots)) {
            return FromConfiguredRoots(configuredRoots);
        }
        if (string.IsNullOrWhiteSpace(workingDirectory)) {
            throw new ArgumentException("An MCP working directory is required.", nameof(workingDirectory));
        }
        return new AgentPathPolicy(new[] { workingDirectory });
    }

    internal string ResolveInput(string path) {
        if (string.IsNullOrWhiteSpace(path)) throw new AgentUsageException("An input path is required.");
        string resolved = OfficeImoToolPathSafety.ResolveExistingLinks(path);
        if (!File.Exists(resolved) && !Directory.Exists(resolved)) {
            throw new FileNotFoundException("Input path '" + resolved + "' does not exist.", resolved);
        }
        EnsureAllowed(resolved);
        return resolved;
    }

    internal string ResolveOutput(string path) {
        if (string.IsNullOrWhiteSpace(path)) throw new AgentUsageException("An output path is required.");
        string resolved = OfficeImoToolPathSafety.ResolveExistingLinks(path);
        EnsureAllowed(resolved);
        return resolved;
    }

    private void EnsureAllowed(string path) {
        if (_allowedRoots.Count == 0) return;
        if (_allowedRoots.Any(root => IsSameOrChildPathExact(root, path))) return;
        throw new UnauthorizedAccessException(
            "Path is outside the configured OfficeIMO MCP roots. Configure " +
            AllowedRootsEnvironmentVariable + " to allow it.");
    }

    private static AgentPathPolicy FromConfiguredRoots(string configuredRoots) {
        string[] roots = configuredRoots.Split(
            Path.PathSeparator,
            StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (roots.Length == 0) {
            throw new AgentUsageException(
                AllowedRootsEnvironmentVariable +
                " must contain at least one directory when it is configured.");
        }
        return new AgentPathPolicy(roots);
    }

    private static bool IsSameOrChildPathExact(string parentPath, string candidatePath) {
        if (string.Equals(parentPath, candidatePath, StringComparison.Ordinal)) return true;
        string parentPrefix = Path.EndsInDirectorySeparator(parentPath)
            ? parentPath
            : parentPath + Path.DirectorySeparatorChar;
        return candidatePath.StartsWith(parentPrefix, StringComparison.Ordinal);
    }
}

internal sealed class AgentUsageException : Exception {
    internal AgentUsageException(string message) : base(message) { }
    internal AgentUsageException(string message, Exception innerException) : base(message, innerException) { }
}
