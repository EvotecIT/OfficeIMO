using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.RealWorldCorpus;

/// <summary>
/// Fetches the pinned third-party test corpora listed in the source manifest. Each source is a Git repository at a
/// fixed commit; only its listed folders are checked out, files are filtered by extension, and the set must match the
/// pinned file count and content digest before anything is copied to the output. The files are never committed here.
/// </summary>
internal static class CorpusSources {
    internal const string ProvenanceFileName = "sources.json";
    internal const string FilesFolderName = "files";

    public static async Task<int> FetchAsync(string[] args) {
        Dictionary<string, string?> values = ParseArguments(args);
        string manifestPath = Path.GetFullPath(Required(values, "--manifest"));
        string outputRoot = Path.GetFullPath(Required(values, "--output"));
        bool printPins = values.ContainsKey("--print-pins");
        CorpusSourceManifest manifest = Load(manifestPath);
        IReadOnlyList<CorpusSource> sources = SelectSources(manifest, values.TryGetValue("--sources", out string? list) ? list : null);

        Directory.CreateDirectory(outputRoot);
        var fetched = new List<CorpusSourceProvenance>();
        bool allVerified = true;
        foreach (CorpusSource source in sources) {
            CorpusSourceProvenance provenance = await FetchSourceAsync(source, outputRoot, printPins).ConfigureAwait(false);
            fetched.Add(provenance);
            allVerified &= provenance.Verified;
            Console.WriteLine($"{source.Id}: {provenance.FileCount} files, {provenance.TotalBytes} bytes, content SHA-256 {provenance.ContentSha256} ({(provenance.Verified ? "verified" : provenance.Status)})");
        }

        string provenancePath = Path.Combine(outputRoot, ProvenanceFileName);
        File.WriteAllText(provenancePath, JsonSerializer.Serialize(fetched, CorpusJson.Options) + Environment.NewLine, new UTF8Encoding(false));
        Console.WriteLine($"Provenance: {provenancePath}");
        return allVerified || printPins ? 0 : 3;
    }

    internal static CorpusSourceManifest Load(string manifestPath) {
        CorpusSourceManifest manifest = JsonSerializer.Deserialize<CorpusSourceManifest>(File.ReadAllText(manifestPath), CorpusJson.Options)
            ?? throw new InvalidDataException("The corpus source manifest is empty.");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (CorpusSource source in manifest.Sources) {
            if (string.IsNullOrWhiteSpace(source.Id) || source.Id.Any(character => !(char.IsAsciiLetterLower(character) || char.IsAsciiDigit(character) || character == '-'))) {
                throw new InvalidDataException("Each corpus source needs a lower-case id of letters, digits and hyphens.");
            }
            if (!ids.Add(source.Id)) throw new InvalidDataException($"The corpus source '{source.Id}' is listed twice.");
            if (!source.Repository.StartsWith("https://", StringComparison.Ordinal)) {
                throw new InvalidDataException($"The corpus source '{source.Id}' must use an https repository URL.");
            }
            if (source.Commit.Length != 40 || source.Commit.Any(character => !Uri.IsHexDigit(character))) {
                throw new InvalidDataException($"The corpus source '{source.Id}' must pin a full 40-character commit.");
            }
            if (source.Paths.Count == 0 || source.Paths.Any(path => path.Length == 0 || path.StartsWith('/') || path.Split('/').Contains(".."))) {
                throw new InvalidDataException($"The corpus source '{source.Id}' needs relative folder paths.");
            }
            if (source.Extensions.Count == 0 || string.IsNullOrWhiteSpace(source.License)) {
                throw new InvalidDataException($"The corpus source '{source.Id}' needs its extensions and license.");
            }
        }
        return manifest;
    }

    private static IReadOnlyList<CorpusSource> SelectSources(CorpusSourceManifest manifest, string? list) {
        if (string.IsNullOrWhiteSpace(list)) return manifest.Sources;
        string[] requested = list.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return requested.Select(id => manifest.Sources.FirstOrDefault(source => source.Id == id)
            ?? throw new ArgumentException($"The manifest has no corpus source '{id}'.")).ToArray();
    }

    private static async Task<CorpusSourceProvenance> FetchSourceAsync(CorpusSource source, string outputRoot, bool printPins) {
        string checkout = Path.Combine(outputRoot, ".checkout-" + source.Id);
        string destination = Path.Combine(outputRoot, FilesFolderName, source.Id);
        DeleteDirectory(checkout);
        DeleteDirectory(destination);
        Directory.CreateDirectory(checkout);
        try {
            await GitAsync(checkout, "init", "--quiet").ConfigureAwait(false);
            await GitAsync(checkout, "remote", "add", "origin", source.Repository).ConfigureAwait(false);
            await GitAsync(checkout, "-c", "protocol.version=2", "fetch", "--quiet", "--depth", "1", "--filter=blob:none", "origin", source.Commit).ConfigureAwait(false);
            var sparse = new List<string> { "sparse-checkout", "set", "--no-cone" };
            sparse.AddRange(source.Paths.Select(path => "/" + path.TrimEnd('/') + "/"));
            await GitAsync(checkout, sparse.ToArray()).ConfigureAwait(false);
            await GitAsync(checkout, "-c", "core.autocrlf=false", "-c", "core.symlinks=false", "checkout", "--quiet", "FETCH_HEAD").ConfigureAwait(false);
            string head = (await GitAsync(checkout, "rev-parse", "HEAD").ConfigureAwait(false)).Trim();
            if (!string.Equals(head, source.Commit, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException($"Checked out {head} instead of the pinned commit {source.Commit}.");
            }

            var extensions = new HashSet<string>(source.Extensions.Select(extension => extension.ToLowerInvariant()), StringComparer.Ordinal);
            var files = new List<(string Relative, string Full)>();
            foreach (string path in source.Paths) {
                string root = Path.Combine(checkout, path.Replace('/', Path.DirectorySeparatorChar));
                if (!Directory.Exists(root)) throw new DirectoryNotFoundException($"The pinned commit has no folder '{path}'.");
                foreach (string file in Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories)) {
                    if (!extensions.Contains(Path.GetExtension(file).ToLowerInvariant())) continue;
                    if ((File.GetAttributes(file) & FileAttributes.ReparsePoint) != 0) continue;
                    files.Add((Path.GetRelativePath(checkout, file).Replace(Path.DirectorySeparatorChar, '/'), file));
                }
            }
            files.Sort((left, right) => string.CompareOrdinal(left.Relative, right.Relative));

            var digest = new StringBuilder();
            long totalBytes = 0;
            foreach ((string relative, string full) in files) {
                byte[] content = await File.ReadAllBytesAsync(full).ConfigureAwait(false);
                totalBytes += content.Length;
                digest.Append(Convert.ToHexString(SHA256.HashData(content)).ToLowerInvariant()).Append(' ').Append(relative).Append('\n');
            }
            string contentSha256 = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(digest.ToString()))).ToLowerInvariant();
            bool verified = files.Count == source.FileCount && string.Equals(contentSha256, source.ContentSha256, StringComparison.OrdinalIgnoreCase);
            var provenance = new CorpusSourceProvenance {
                Id = source.Id,
                Name = source.Name,
                Repository = source.Repository,
                Commit = source.Commit,
                Paths = source.Paths,
                License = source.License,
                LicenseUrl = source.LicenseUrl,
                FileCount = files.Count,
                TotalBytes = totalBytes,
                ContentSha256 = contentSha256,
                Verified = verified,
                Status = verified ? "verified" : $"content mismatch: expected {source.FileCount} files with SHA-256 {source.ContentSha256}"
            };
            if (verified || printPins) {
                foreach ((string relative, string full) in files) {
                    string target = Path.Combine(destination, relative.Replace('/', Path.DirectorySeparatorChar));
                    Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                    File.Copy(full, target);
                }
            }
            return provenance;
        } finally {
            DeleteDirectory(checkout);
        }
    }

    private static async Task<string> GitAsync(string workingDirectory, params string[] arguments) {
        var startInfo = new ProcessStartInfo("git") {
            WorkingDirectory = workingDirectory,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        foreach (string argument in arguments) startInfo.ArgumentList.Add(argument);
        startInfo.Environment["GIT_TERMINAL_PROMPT"] = "0";
        using Process process = Process.Start(startInfo) ?? throw new InvalidOperationException("Unable to start git.");
        Task<string> output = process.StandardOutput.ReadToEndAsync();
        Task<string> error = process.StandardError.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(20));
        try {
            await process.WaitForExitAsync(timeout.Token).ConfigureAwait(false);
        } catch (OperationCanceledException) {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException($"git {arguments[0]} did not finish within 20 minutes.");
        }
        string stdout = await output.ConfigureAwait(false);
        string stderr = await error.ConfigureAwait(false);
        if (process.ExitCode != 0) {
            throw new InvalidOperationException($"git {string.Join(' ', arguments)} failed with exit code {process.ExitCode}: {stderr.Trim()}");
        }
        return stdout;
    }

    private static void DeleteDirectory(string path) {
        if (!Directory.Exists(path)) return;
        foreach (string file in Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories)) {
            File.SetAttributes(file, FileAttributes.Normal);
        }
        Directory.Delete(path, recursive: true);
    }

    private static Dictionary<string, string?> ParseArguments(string[] args) {
        var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        for (int index = 0; index < args.Length; index++) {
            string name = args[index];
            switch (name) {
                case "--print-pins":
                    values[name] = null;
                    break;
                case "--manifest":
                case "--output":
                case "--sources":
                    if (index + 1 >= args.Length) throw new ArgumentException($"{name} requires a value.");
                    values[name] = args[++index];
                    break;
                default:
                    throw new ArgumentException($"Unknown option '{name}'.");
            }
        }
        return values;
    }

    private static string Required(IReadOnlyDictionary<string, string?> values, string name) =>
        values.TryGetValue(name, out string? value) && value != null ? value : throw new ArgumentException($"{name} is required.");
}

internal sealed class CorpusSourceManifest {
    public int SchemaVersion { get; set; }
    public IReadOnlyList<CorpusSource> Sources { get; set; } = Array.Empty<CorpusSource>();
}

internal sealed class CorpusSource {
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Repository { get; set; } = string.Empty;
    public string Commit { get; set; } = string.Empty;
    public IReadOnlyList<string> Paths { get; set; } = Array.Empty<string>();
    public IReadOnlyList<string> Extensions { get; set; } = Array.Empty<string>();
    public string License { get; set; } = string.Empty;
    public string? LicenseUrl { get; set; }
    public int FileCount { get; set; }
    public string ContentSha256 { get; set; } = string.Empty;
}

internal sealed class CorpusSourceProvenance {
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Repository { get; set; } = string.Empty;
    public string Commit { get; set; } = string.Empty;
    public IReadOnlyList<string> Paths { get; set; } = Array.Empty<string>();
    public string License { get; set; } = string.Empty;
    public string? LicenseUrl { get; set; }
    public int FileCount { get; set; }
    public long TotalBytes { get; set; }
    public string ContentSha256 { get; set; } = string.Empty;
    public bool Verified { get; set; }
    public string Status { get; set; } = string.Empty;
}
