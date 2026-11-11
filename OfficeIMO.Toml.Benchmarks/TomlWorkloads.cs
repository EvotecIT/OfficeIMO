using System.Text;
using OfficeIMO.Tests;

namespace OfficeIMO.Toml.Benchmarks;

/// <summary>Benchmark inputs, all from the pinned real-world corpus.</summary>
public enum TomlWorkload {
    /// <summary>PyPA's commented sample pyproject.toml (7 KB).</summary>
    SmallConfig,
    /// <summary>uv's workspace Cargo.toml with hundreds of inline-table dependencies (25 KB).</summary>
    CargoWorkspace,
    /// <summary>uv's uv.lock: thousands of [[package]] tables and inline tables (350 KB).</summary>
    LockFile,
    /// <summary>
    /// Every corpus file once, except uv's ruff.toml: it uses a TOML 1.1 multi-line inline table, which Tomlet, a TOML 1.0
    /// reader, rejects.
    /// </summary>
    Corpus
}

public static class TomlWorkloads {
    private static readonly Dictionary<TomlWorkload, string[]> Cache = new();

    public static string[] Get(TomlWorkload workload) {
        lock (Cache) {
            if (!Cache.TryGetValue(workload, out string[]? inputs)) {
                inputs = workload switch {
                    TomlWorkload.SmallConfig => new[] { Read("Python/sampleproject-pyproject.toml") },
                    TomlWorkload.CargoWorkspace => new[] { Read("Cargo/uv-Cargo.toml") },
                    TomlWorkload.LockFile => new[] { Read("Uv/uv-uv.lock") },
                    _ => Directory.GetFiles(CorpusRoot, "*.*", SearchOption.AllDirectories)
                        .Where(path => path.EndsWith(".toml", StringComparison.Ordinal) || path.EndsWith(".lock", StringComparison.Ordinal))
                        .Where(path => !path.EndsWith("uv-ruff.toml", StringComparison.Ordinal))
                        .OrderBy(path => path, StringComparer.Ordinal)
                        .Select(path => File.ReadAllText(path, Encoding.UTF8))
                        .ToArray()
                };
                Cache[workload] = inputs;
            }
            return inputs;
        }
    }

    private static string CorpusRoot => Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Toml", "Corpus");

    private static string Read(string path) => File.ReadAllText(Path.Combine(CorpusRoot, path), Encoding.UTF8);
}
