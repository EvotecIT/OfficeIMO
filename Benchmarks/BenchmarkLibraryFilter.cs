using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Filters;

namespace OfficeIMO.Benchmarks;

/// <summary>
/// Leaves out BenchmarkDotNet cases of restricted libraries the caller has not enabled. A case names its library with
/// a <c>library:&lt;id&gt;</c> category; only libraries that Build/benchmark-libraries.json marks as anything other than
/// <c>open</c> are left out, so open libraries (OfficeIMO included) keep running when tagged. The policy is embedded at
/// build time by Benchmarks/BenchmarkLibraryFilter.props.
/// </summary>
internal static class BenchmarkLibraryFilter {
    internal const string PolicyResourceName = "OfficeIMO.Benchmarks.benchmark-libraries.json";

    private static readonly Lazy<IReadOnlyDictionary<string, BenchmarkPolicyLibrary>> Policy = new(LoadPolicy);

    internal static ManualConfig AddRestrictedLibraryFilter(this ManualConfig config) {
        // Read the policy now so a missing embed fails before any benchmark is selected.
        _ = Policy.Value;
        return config.AddFilter(new SimpleFilter(benchmarkCase => ShouldRun(benchmarkCase.Descriptor.Categories)));
    }

    /// <summary>
    /// Whether a case with <paramref name="categories"/> runs: every <c>library:</c> category must name an open library
    /// or a restricted one listed in <see cref="BenchmarkLibrarySelection.EnabledLibrariesVariable"/>.
    /// </summary>
    internal static bool ShouldRun(IEnumerable<string> categories) {
        foreach (string category in categories) {
            if (!category.StartsWith(BenchmarkLibrarySelection.CategoryPrefix, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            string name = category.Substring(BenchmarkLibrarySelection.CategoryPrefix.Length).Trim();
            if (!Policy.Value.TryGetValue(name, out BenchmarkPolicyLibrary? library)) {
                throw new InvalidOperationException(
                    $"Benchmark category '{category}' names a library that Build/benchmark-libraries.json does not list.");
            }

            if (!string.Equals(library.Restriction, "open", StringComparison.Ordinal) &&
                !BenchmarkLibrarySelection.IsEnabled(library.Id) && !BenchmarkLibrarySelection.IsEnabled(name)) {
                return false;
            }
        }

        return true;
    }

    private static IReadOnlyDictionary<string, BenchmarkPolicyLibrary> LoadPolicy() {
        using Stream? stream = typeof(BenchmarkLibraryFilter).Assembly.GetManifestResourceStream(PolicyResourceName);
        if (stream is null) {
            throw new InvalidOperationException(
                $"The benchmark library policy is not embedded in {typeof(BenchmarkLibraryFilter).Assembly.GetName().Name}. " +
                "Import Benchmarks/BenchmarkLibraryFilter.props in the project that compiles BenchmarkLibraryFilter.cs.");
        }

        BenchmarkPolicyDocument? document;
        try {
            document = (BenchmarkPolicyDocument?)new DataContractJsonSerializer(typeof(BenchmarkPolicyDocument)).ReadObject(stream);
        } catch (SerializationException exception) {
            throw new InvalidOperationException("The embedded benchmark library policy is not valid JSON.", exception);
        }

        if (document?.Libraries is not { Length: > 0 }) {
            throw new InvalidOperationException("The embedded benchmark library policy lists no libraries.");
        }

        var libraries = new Dictionary<string, BenchmarkPolicyLibrary>(StringComparer.OrdinalIgnoreCase);
        foreach (BenchmarkPolicyLibrary library in document.Libraries) {
            if (string.IsNullOrWhiteSpace(library.Id) || string.IsNullOrWhiteSpace(library.Restriction)) {
                throw new InvalidOperationException("Every library in the embedded benchmark library policy needs an id and a restriction.");
            }

            IEnumerable<string?> names = new[] { library.Id, library.Name }
                .Concat(library.Aliases ?? [])
                .Concat((library.Declare ?? []).Select(item => item.Id));
            foreach (string name in names.Where(name => !string.IsNullOrWhiteSpace(name)).Select(name => name!).Distinct(StringComparer.OrdinalIgnoreCase)) {
                if (libraries.TryGetValue(name, out BenchmarkPolicyLibrary? existing) && !ReferenceEquals(existing, library)) {
                    throw new InvalidOperationException(
                        $"The benchmark library policy names both '{existing.Id}' and '{library.Id}' as '{name}'.");
                }

                libraries[name] = library;
            }
        }

        return libraries;
    }

    [DataContract]
    private sealed class BenchmarkPolicyDocument {
        [DataMember(Name = "libraries")] public BenchmarkPolicyLibrary[]? Libraries { get; set; }
    }

    [DataContract]
    private sealed class BenchmarkPolicyLibrary {
        [DataMember(Name = "id")] public string Id { get; set; } = string.Empty;
        [DataMember(Name = "name")] public string? Name { get; set; }
        [DataMember(Name = "aliases")] public string[]? Aliases { get; set; }
        [DataMember(Name = "restriction")] public string Restriction { get; set; } = string.Empty;
        [DataMember(Name = "declare")] public BenchmarkPolicyDeclaration[]? Declare { get; set; }
    }

    [DataContract]
    private sealed class BenchmarkPolicyDeclaration {
        [DataMember(Name = "id")] public string? Id { get; set; }
    }
}
