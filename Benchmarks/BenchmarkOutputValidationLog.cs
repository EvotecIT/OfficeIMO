using System;
using System.IO;
using System.Text;

namespace OfficeIMO.Benchmarks;

/// <summary>
/// Records the output validation a comparison performs before timing, so the benchmark evidence catalog can tell a
/// validated result from an unchecked one. The library comparison runner names the file in
/// <see cref="PathVariable"/>; a run without it records nothing. Each call appends one JSON object
/// (<c>library</c>, <c>case</c>, <c>status</c>, <c>detail</c>) on its own line.
/// </summary>
internal static class BenchmarkOutputValidationLog {
    internal const string PathVariable = "OFFICEIMO_BENCHMARK_VALIDATION_PATH";

    private static readonly object Gate = new();

    /// <summary>
    /// Records that <paramref name="library"/>'s output matched the expected result. <paramref name="library"/> is
    /// the id in Build/benchmark-declarations.json; an empty <paramref name="caseName"/> covers every case of it.
    /// </summary>
    internal static void Passed(string library, string detail, string caseName = "") => Write(library, caseName, "passed", detail);

    /// <summary>Records that <paramref name="library"/>'s output did not match the expected result.</summary>
    internal static void Failed(string library, string detail, string caseName = "") => Write(library, caseName, "failed", detail);

    private static void Write(string library, string caseName, string status, string detail) {
        if (string.IsNullOrWhiteSpace(library)) {
            throw new ArgumentException("A validation needs the library it covers.", nameof(library));
        }

        string? path = Environment.GetEnvironmentVariable(PathVariable);
        if (string.IsNullOrWhiteSpace(path)) {
            return;
        }

        string line = "{\"library\":" + Quote(library) + ",\"case\":" + Quote(caseName) + ",\"status\":" + Quote(status) +
            ",\"detail\":" + Quote(detail) + "}" + "\n";
        byte[] bytes = new UTF8Encoding(false).GetBytes(line);
        lock (Gate) {
            // Benchmark processes run one after another, but a process may record from several threads.
            using var stream = new FileStream(path!, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
            stream.Write(bytes, 0, bytes.Length);
        }
    }

    private static string Quote(string value) {
        var builder = new StringBuilder(value.Length + 2).Append('"');
        foreach (char character in value) {
            switch (character) {
                case '"': builder.Append("\\\""); break;
                case '\\': builder.Append("\\\\"); break;
                case '\n': builder.Append("\\n"); break;
                case '\r': builder.Append("\\r"); break;
                case '\t': builder.Append("\\t"); break;
                default:
                    if (character < ' ') {
                        builder.Append("\\u").Append(((int)character).ToString("x4", System.Globalization.CultureInfo.InvariantCulture));
                    } else {
                        builder.Append(character);
                    }
                    break;
            }
        }
        return builder.Append('"').ToString();
    }
}
