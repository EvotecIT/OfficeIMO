using System.Globalization;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using OfficeIMO.Tests;

namespace OfficeIMO.Toml.Benchmarks.LargeCorpus;

/// <summary>One pinned file of the large corpus.</summary>
internal sealed class TomlCorpusFile {
    public TomlCorpusFile(string source, string category, string path, string url, string sha256, long bytes, string fullPath) {
        Source = source;
        Category = category;
        Path = path;
        Url = url;
        Sha256 = sha256;
        Bytes = bytes;
        FullPath = fullPath;
    }

    public string Source { get; }

    public string Category { get; }

    /// <summary>Path inside the source repository.</summary>
    public string Path { get; }

    public string Url { get; }

    public string Sha256 { get; }

    public long Bytes { get; }

    public string FullPath { get; }

    public string Name => Source + ":" + Path;

    /// <summary>Reads the cached file and checks its pinned SHA-256.</summary>
    public byte[] Read() {
        byte[] bytes = File.ReadAllBytes(FullPath);
        if (!string.Equals(TomlLargeCorpus.Sha256(bytes), Sha256, StringComparison.Ordinal)) throw new InvalidDataException(Name + " does not match its pinned SHA-256; run --fetch-corpus.");
        return bytes;
    }
}

/// <summary>
/// The opt-in large corpus pinned in <c>OfficeIMO.TestAssets/Toml/LargeCorpus/sources.json</c> by source URL and SHA-256.
/// The cache is <c>--cache</c>, else <c>OFFICEIMO_TOML_CORPUS_CACHE</c>, else <c>toml-corpus</c> under the temporary
/// folder; nothing from it is committed.
/// </summary>
internal static class TomlLargeCorpus {
    private static readonly JsonSerializerOptions Indented = new() { WriteIndented = true, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    public static string ManifestPath => System.IO.Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Toml", "LargeCorpus", "sources.json");

    public static string ResolveRoot(string? option) {
        string? root = option ?? Environment.GetEnvironmentVariable("OFFICEIMO_TOML_CORPUS_CACHE");
        return System.IO.Path.GetFullPath(string.IsNullOrWhiteSpace(root) ? System.IO.Path.Combine(System.IO.Path.GetTempPath(), "toml-corpus") : root);
    }

    /// <summary>Lists every pinned file with its cache location.</summary>
    public static List<TomlCorpusFile> Files(string root) {
        JsonNode manifest = JsonNode.Parse(File.ReadAllText(ManifestPath))!;
        var files = new List<TomlCorpusFile>();
        foreach (JsonNode? source in manifest["sources"]!.AsArray()) {
            string id = (string)source!["id"]!;
            string category = (string)source["category"]!;
            string raw = RawBase(source);
            foreach (JsonNode? file in source["files"]?.AsArray() ?? new JsonArray()) {
                string path = (string)file!["path"]!;
                string fullPath = System.IO.Path.Combine(root, "sources", id, path.Replace('/', System.IO.Path.DirectorySeparatorChar));
                files.Add(new TomlCorpusFile(id, category, path, raw + EscapePath(path), (string)file["sha256"]!, (long)file["bytes"]!, fullPath));
            }
        }
        return files;
    }

    /// <summary><c>--fetch-corpus</c>: downloads missing or changed files and checks every SHA-256.</summary>
    public static int Fetch(string[] args) {
        string root = ResolveRoot(Option(args, "--cache"));
        List<TomlCorpusFile> files = Files(root);
        using HttpClient http = CreateClient(api: false);
        int downloaded = 0, failed = 0;
        Parallel.ForEach(files, new ParallelOptions { MaxDegreeOfParallelism = 12 }, file => {
            if (File.Exists(file.FullPath) && Sha256(File.ReadAllBytes(file.FullPath)) == file.Sha256) return;
            try {
                byte[] bytes = Download(http, file.Url);
                if (Sha256(bytes) != file.Sha256) throw new InvalidDataException("SHA-256 mismatch");
                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(file.FullPath)!);
                File.WriteAllBytes(file.FullPath, bytes);
                Interlocked.Increment(ref downloaded);
            } catch (Exception ex) {
                Interlocked.Increment(ref failed);
                Console.Error.WriteLine(file.Name + ": " + ex.Message);
            }
        });
        Console.WriteLine(Format(files.Count) + " files pinned, " + Format(downloaded) + " downloaded, " + Format(failed) + " failed; cache " + root);
        return failed == 0 ? 0 : 1;
    }

    /// <summary>
    /// <c>--refresh-corpus [--source id]</c>: lists each source's trees at its pinned commit, applies the globs and the
    /// sample size, downloads the selected files into the cache and records their SHA-256 and size in the manifest.
    /// Set <c>GITHUB_TOKEN</c> to raise the API rate limit.
    /// </summary>
    public static int Refresh(string[] args) {
        string root = ResolveRoot(Option(args, "--cache"));
        string? only = Option(args, "--source");
        JsonNode manifest = JsonNode.Parse(File.ReadAllText(ManifestPath))!;
        using HttpClient api = CreateClient(api: true);
        using HttpClient raw = CreateClient(api: false);
        int failures = 0;
        foreach (JsonNode? node in manifest["sources"]!.AsArray()) {
            JsonObject source = node!.AsObject();
            string id = (string)source["id"]!;
            if (only is not null && only != id) continue;
            string repository = (string)source["repository"]!;
            string slug = new Uri(repository).AbsolutePath.Trim('/');
            string commit = (string)source["commit"]!;
            Regex include = Glob(source["include"]!.AsArray().Select(g => (string)g!));
            var paths = new SortedSet<string>(StringComparer.Ordinal);
            foreach (JsonNode? r in source["roots"]!.AsArray()) {
                string prefix = (string)r!;
                foreach (string path in ListTree(api, slug, commit + (prefix.Length == 0 ? string.Empty : ":" + prefix), prefix.Length == 0 ? string.Empty : prefix + "/")) {
                    if (include.IsMatch(path)) paths.Add(path);
                }
            }
            List<string> selected = Sample(paths.ToList(), source["max"] is JsonNode max ? (int)max : int.MaxValue);
            string rawBase = RawBase(source);
            var files = new JsonObject[selected.Count];
            Parallel.For(0, selected.Count, new ParallelOptions { MaxDegreeOfParallelism = 12 }, i => {
                string path = selected[i];
                try {
                    byte[] bytes = Download(raw, rawBase + EscapePath(path));
                    string fullPath = System.IO.Path.Combine(root, "sources", id, path.Replace('/', System.IO.Path.DirectorySeparatorChar));
                    Directory.CreateDirectory(System.IO.Path.GetDirectoryName(fullPath)!);
                    File.WriteAllBytes(fullPath, bytes);
                    files[i] = new JsonObject { ["path"] = path, ["sha256"] = Sha256(bytes), ["bytes"] = bytes.LongLength };
                } catch (Exception ex) {
                    Interlocked.Increment(ref failures);
                    Console.Error.WriteLine(id + ":" + path + ": " + ex.Message);
                }
            });
            source["files"] = new JsonArray(files.Where(f => f is not null).Select(f => (JsonNode)f).ToArray());
            Console.WriteLine(id + ": " + paths.Count + " matches, " + files.Count(f => f is not null) + " pinned");
        }
        // One line per pinned file keeps the manifest compact and its diffs readable.
        string json = FileEntry.Replace(manifest.ToJsonString(Indented).Replace("\r\n", "\n"), "{ $1, $2, $3 }");
        File.WriteAllText(ManifestPath, json + "\n", new UTF8Encoding(false));
        return failures == 0 ? 0 : 1;
    }

    private static readonly Regex FileEntry = new(@"\{\n\s*(""path"": ""[^""\n]*""),\n\s*(""sha256"": ""[0-9a-f]+""),\n\s*(""bytes"": \d+)\n\s*\}", RegexOptions.CultureInvariant);

    /// <summary>Lists blob paths under a tree, descending into subtrees one at a time when GitHub truncates the listing.</summary>
    private static IEnumerable<string> ListTree(HttpClient api, string slug, string tree, string prefix) {
        JsonNode listing = JsonNode.Parse(Get(api, "https://api.github.com/repos/" + slug + "/git/trees/" + tree + "?recursive=1"))!;
        bool truncated = (bool?)listing["truncated"] ?? false;
        if (!truncated) {
            foreach (JsonNode? entry in listing["tree"]!.AsArray()) {
                if ((string)entry!["type"]! == "blob") yield return prefix + (string)entry["path"]!;
            }
            yield break;
        }
        JsonNode level = JsonNode.Parse(Get(api, "https://api.github.com/repos/" + slug + "/git/trees/" + tree))!;
        foreach (JsonNode? entry in level["tree"]!.AsArray()) {
            string name = (string)entry!["path"]!;
            if ((string)entry["type"]! == "blob") yield return prefix + name;
            else if ((string)entry["type"]! == "tree") {
                foreach (string path in ListTree(api, slug, (string)entry["sha"]!, prefix + name + "/")) yield return path;
            }
        }
    }

    private static List<string> Sample(List<string> sorted, int max) {
        if (sorted.Count <= max) return sorted;
        var picked = new List<string>(max);
        for (int i = 0; i < max; i++) picked.Add(sorted[(int)((long)i * sorted.Count / max)]);
        return picked;
    }

    /// <summary>Translates globs ('**' spans folders, '*' and '?' stay in one, '{a,b}' lists alternatives) to one regex.</summary>
    private static Regex Glob(IEnumerable<string> globs) {
        var pattern = new StringBuilder("^(?:");
        bool first = true;
        foreach (string glob in globs) {
            if (!first) pattern.Append('|');
            first = false;
            for (int i = 0; i < glob.Length; i++) {
                char ch = glob[i];
                if (ch == '*' && i + 1 < glob.Length && glob[i + 1] == '*') {
                    bool slash = i + 2 < glob.Length && glob[i + 2] == '/';
                    pattern.Append(slash ? "(?:.*/)?" : ".*");
                    i += slash ? 2 : 1;
                } else if (ch == '*') {
                    pattern.Append("[^/]*");
                } else if (ch == '?') {
                    pattern.Append("[^/]");
                } else if (ch == '{') {
                    pattern.Append("(?:");
                } else if (ch == '}') {
                    pattern.Append(')');
                } else if (ch == ',') {
                    pattern.Append('|');
                } else {
                    pattern.Append(Regex.Escape(ch.ToString()));
                }
            }
        }
        return new Regex(pattern.Append(")$").ToString(), RegexOptions.CultureInvariant);
    }

    private static string RawBase(JsonNode source) {
        string slug = new Uri((string)source["repository"]!).AbsolutePath.Trim('/');
        return "https://raw.githubusercontent.com/" + slug + "/" + (string)source["commit"]! + "/";
    }

    private static string EscapePath(string path) => string.Join("/", path.Split('/').Select(Uri.EscapeDataString));

    private static HttpClient CreateClient(bool api) {
        var http = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };
        http.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("OfficeIMO-Toml-Corpus", "1.0"));
        string? token = Environment.GetEnvironmentVariable("GITHUB_TOKEN");
        if (api && !string.IsNullOrEmpty(token)) http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        return http;
    }

    private static string Get(HttpClient http, string url) => Encoding.UTF8.GetString(Download(http, url));

    /// <summary>Downloads with a few retries; GitHub resets connections under load.</summary>
    private static byte[] Download(HttpClient http, string url) {
        for (int attempt = 1; ; attempt++) {
            try {
                using HttpResponseMessage response = http.GetAsync(url).GetAwaiter().GetResult();
                response.EnsureSuccessStatusCode();
                return response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
            } catch (Exception) when (attempt < 5) {
                Thread.Sleep(1000 * attempt);
            }
        }
    }

    public static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    public static string? Option(string[] args, string name) {
        int index = Array.IndexOf(args, name);
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }

    public static string Format(long value) => value.ToString("N0", CultureInfo.InvariantCulture);
}
