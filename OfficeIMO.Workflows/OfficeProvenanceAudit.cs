using System.IO.Enumeration;
using OfficeIMO.Internal;
using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows;

/// <summary>Read-only bounded file and directory audit settings.</summary>
public sealed class OfficeProvenanceAuditRequest {
    /// <summary>Explicit files and directories to assess. Unsupported explicit files produce failed items.</summary>
    public IReadOnlyList<string> Inputs { get; init; } = Array.Empty<string>();
    /// <summary>Whether directory discovery descends into subdirectories.</summary>
    public bool Recursive { get; init; } = true;
    /// <summary>Optional wildcards matched against slash-separated paths relative to each directory root.</summary>
    public IReadOnlyList<string> Include { get; init; } = Array.Empty<string>();
    /// <summary>Wildcards excluded from discovery. Explicit file inputs also observe these patterns.</summary>
    public IReadOnlyList<string> Exclude { get; init; } = Array.Empty<string>();
    /// <summary>Maximum unique assessed inputs, from 1 to 10,000. Exceeding the bound fails discovery.</summary>
    public int MaximumItems { get; init; } = 256;
    /// <summary>Maximum filesystem entries visited, including unsupported and excluded files.</summary>
    public int MaximumVisitedEntries { get; init; } = 100_000;
    /// <summary>Maximum bytes per input.</summary>
    public long MaximumInputBytes { get; init; } = 256L * 1024 * 1024;
    /// <summary>Whether text-like files receive Unicode inspection.</summary>
    public bool InspectTextIntegrity { get; init; } = true;
    /// <summary>Whether qualified owners inspect embedded assets.</summary>
    public bool ProcessEmbeddedAssets { get; init; } = true;
}

/// <summary>Canonical read-only discovery and assessment shared by automation and interactive hosts.</summary>
public static class OfficeProvenanceAudit {
    /// <summary>Discovers a deterministic, bounded set. Directory symlinks and generated/VCS directories are not followed.</summary>
    public static IReadOnlyList<string> Discover(OfficeProvenanceAuditRequest request, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        if (request.Inputs.Count == 0) throw new ArgumentException("At least one audit input is required.");
        if (request.MaximumItems is < 1 or > 10_000 || request.MaximumVisitedEntries < 1 || request.MaximumInputBytes < 1)
            throw new ArgumentOutOfRangeException(nameof(request));
        var paths = new SortedSet<string>(OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
        int visited = 0;
        foreach (string input in request.Inputs) {
            cancellationToken.ThrowIfCancellationRequested();
            string root = Path.GetFullPath(input);
            if (!Directory.Exists(root)) { Add(root, Path.GetFileName(root), null); continue; }
            if ((File.GetAttributes(root) & FileAttributes.ReparsePoint) != 0)
                throw new IOException("Audit directory roots cannot be symbolic links: " + root);
            string physicalRoot = OfficePathIdentity.ResolvePhysicalPath(root);
            var pending = new Stack<(string Path, string Identity)>();
            pending.Push((root, OfficePathIdentity.GetPhysicalIdentityKey(root)));
            while (pending.Count != 0) {
                cancellationToken.ThrowIfCancellationRequested();
                (string directory, string expectedIdentity) = pending.Pop();
                using var handle = OfficePathIdentity.OpenDirectoryForIdentity(directory, out string openedDirectory);
                if (!string.Equals(OfficePathIdentity.GetPhysicalIdentityKey(directory), expectedIdentity, StringComparison.Ordinal) ||
                    !OfficePathIdentity.IsSameOrDescendant(openedDirectory, physicalRoot))
                    throw new InvalidDataException("An audit directory changed during discovery.");
                foreach (string entry in Directory.EnumerateFileSystemEntries(directory)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (++visited > request.MaximumVisitedEntries) throw new InvalidDataException("Audit discovery exceeded the visited-entry limit.");
                    FileAttributes attributes = File.GetAttributes(entry);
                    if ((attributes & FileAttributes.ReparsePoint) != 0) continue;
                    string relative = Path.GetRelativePath(root, entry).Replace('\\', '/');
                    if ((attributes & FileAttributes.Directory) != 0) {
                        string name = Path.GetFileName(entry);
                        if (request.Recursive && name is not ".git" and not "bin" and not "obj" and not "node_modules" && !Excluded(relative))
                            pending.Push((entry, OfficePathIdentity.GetPhysicalIdentityKey(entry)));
                    } else if (OfficeProvenanceWorkflowCatalog.FindByPath(entry) != null) Add(entry, relative, physicalRoot);
                }
                OfficePathIdentity.EnsurePathMatchesOpenedDirectory(directory, handle);
            }
        }
        if (paths.Count == 0) throw new InvalidDataException("No eligible files matched this audit. No assessment was performed.");
        return paths.ToArray();

        bool Excluded(string path) => request.Exclude.Any(pattern => Matches(pattern, path));
        void Add(string path, string relative, string? physicalRoot) {
            if (Excluded(relative) || request.Include.Count != 0 && !request.Include.Any(pattern => Matches(pattern, relative))) return;
            if (physicalRoot != null) using (OfficePathIdentity.OpenRegularFileForRead(path, physicalRoot, 81920)) { }
            paths.Add(path);
            if (paths.Count > request.MaximumItems) throw new InvalidDataException("Audit discovery exceeded the item limit. Narrow the selection or increase MaximumItems.");
        }
    }
    /// <summary>Assesses discovered inputs without publishing or modifying assets.</summary>
    public static async Task<IReadOnlyList<OfficeProvenanceWorkflowResult>> RunAsync(OfficeProvenanceAuditRequest request,
        IOfficeProvenanceWorkflowRunner? runner = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        var roots = request.Inputs.Select(input => {
            string selected = Path.GetFullPath(input);
            string directory = Directory.Exists(selected) ? selected : Path.GetDirectoryName(selected)!;
            string? physical = Directory.Exists(directory) ? OfficePathIdentity.ResolvePhysicalPath(directory) : null;
            return (Selected: selected, IsDirectory: Directory.Exists(selected),
                Physical: physical, Identity: physical == null ? null : OfficePathIdentity.GetPhysicalIdentityKey(physical));
        }).ToArray();
        IReadOnlyList<string> paths = Discover(request, cancellationToken);
        runner ??= new OfficeWorkflowRunner();
        return await runner.RunProvenanceBatchAsync(paths.Select(path => {
            var selectedRoot = roots.Where(root => root.IsDirectory
                    ? IsSelectedPath(path, root.Selected)
                    : string.Equals(path, root.Selected, OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
                .OrderByDescending(root => root.Selected.Length).FirstOrDefault();
            if (selectedRoot.Physical == null && File.Exists(path))
                throw new InvalidDataException("An audit input has no selected source root.");
            var item = new OfficeProvenanceWorkflowRequest { InputPath = path, Operation = OfficeProvenanceWorkflowOperation.Assess,
                Limits = new OfficeWorkflowLimits { MaximumInputBytes = request.MaximumInputBytes },
                AuditRootPhysicalPath = selectedRoot.Physical, AuditRootIdentity = selectedRoot.Identity };
            item.Assessment.Structural.MaxAssetBytes = Math.Min(request.MaximumInputBytes, int.MaxValue);
            item.Assessment.TextIntegrity.MaxEncodedBytes = Math.Min(request.MaximumInputBytes, int.MaxValue);
            item.Assessment.Structural.ProcessEmbeddedAssets = request.ProcessEmbeddedAssets;
            item.Assessment.InspectTextIntegrity = request.InspectTextIntegrity;
            return item;
        }), new OfficeProvenanceWorkflowBatchOptions { MaximumRequests = request.MaximumItems },
            cancellationToken: cancellationToken).ConfigureAwait(false);
    }
    /// <summary>Whether the selected evidence policy has actionable findings. Execution failures must be handled separately.</summary>
    public static bool HasFindings(OfficeProvenanceWorkflowResult result, bool carriers = false, bool dangerousText = true) =>
        carriers && (result.Assessment?.Structural ?? result.Inspection)?.Evidence.Count > 0 ||
        dangerousText && result.Assessment?.TextIntegrity?.HasPotentiallyDangerousFindings == true;
    private static bool Matches(string pattern, string path) => FileSystemName.MatchesSimpleExpression(pattern.Replace('\\', '/'), path, OperatingSystem.IsWindows());
    private static bool IsSelectedPath(string path, string root) {
        string relative = Path.GetRelativePath(root, path);
        return relative == "." || !Path.IsPathRooted(relative) && relative != ".." &&
            !relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal);
    }
}
