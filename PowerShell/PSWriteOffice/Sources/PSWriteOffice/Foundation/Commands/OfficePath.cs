using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;

namespace PSWriteOffice.Foundation;

/// <summary>Resolves filesystem provider paths without interpreting literal bracket characters as wildcards.</summary>
public static class OfficePath {
    /// <summary>Compares physical filesystem identities, including directory aliases and links.</summary>
    public static bool AreEquivalent(string left, string right) => OfficeIMO.Internal.OfficePathIdentity.AreEquivalent(left, right);

    /// <summary>Reads literal input paths when object pipeline binding precedes LiteralPath property binding.</summary>
    internal static IReadOnlyList<string>? GetPipelineLiteralPaths(object input) {
        if (OfficeArgumentValue.Unwrap(input) is System.IO.FileInfo file) return new[] { file.FullName };
        // Keep the original wrapper: unwrapping PSCustomObject loses its per-instance properties.
        PSObject value = PSObject.AsPSObject(input);
        foreach (string name in new[] { "LiteralPath", "PSPath", "FullName" }) {
            PSPropertyInfo? property = value.Properties[name];
            if (property == null) continue;
            string[] paths = LanguagePrimitives.ConvertTo<string[]>(property.Value);
            if (paths == null || paths.Length == 0)
                throw new PSArgumentException("LiteralPath requires at least one path.", "LiteralPath");
            return paths;
        }
        return null;
    }

    /// <summary>Resolves existing input items through the current PowerShell provider; an optional reporter retains valid wildcard matches when other items are rejected.</summary>
    public static IReadOnlyList<string> ResolveInput(SessionState session, string path, bool literal, bool allowDirectory = false, Action<Exception, string>? reportItemError = null) {
        if (session == null) throw new ArgumentNullException(nameof(session));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("An input path is required.", nameof(path));
        var result = new List<string>();
        if (literal) {
            string full = ResolveProviderPath(session, path);
            CheckInput(full, allowDirectory);
            result.Add(full);
        } else {
            var paths = session.Path.GetResolvedProviderPathFromPSPath(path, out var provider);
            CheckProvider(provider);
            foreach (string full in paths) {
                try { CheckInput(full, allowDirectory); result.Add(full); }
                catch (FileNotFoundException error) when (reportItemError != null) { reportItemError(error, full); }
            }
            if (paths.Count == 0) throw new FileNotFoundException("No input files matched the path.", path);
        }
        return result;
    }
    /// <summary>Resolves a single output file or directory, including paths that do not exist.</summary>
    public static string ResolveOutput(SessionState session, string path, bool directory) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("An output path is required.", nameof(path));
        string full = ResolveProviderPath(session, path);
        if (!directory && (Directory.Exists(full) || path.EndsWith("/", StringComparison.Ordinal) || path.EndsWith("\\", StringComparison.Ordinal)))
            throw new ArgumentException("OutputPath names a file. Use OutputDirectory for a folder.", nameof(path));
        if (directory && File.Exists(full)) throw new ArgumentException("OutputDirectory names a folder, not an existing file.", nameof(path));
        return full;
    }
    /// <summary>Validates the New/Force or Export/Convert/NoClobber overwrite policy before an effect.</summary>
    public static void ValidateDestination(string path, bool refuseExisting) {
        if (Directory.Exists(path)) throw new ArgumentException("The output file is an existing directory.", nameof(path));
        if (refuseExisting && File.Exists(path)) throw new IOException("The output already exists: " + path);
    }
    private static string ResolveProviderPath(SessionState session, string path) {
        string full = session.Path.GetUnresolvedProviderPathFromPSPath(path, out var provider, out _);
        CheckProvider(provider);
        return full;
    }
    private static void CheckProvider(ProviderInfo provider) {
        if (!provider.Name.Equals("FileSystem", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Only FileSystem provider paths are supported.");
    }
    private static void CheckInput(string full, bool allowDirectory) {
        if (!File.Exists(full) && !(allowDirectory && Directory.Exists(full))) throw new FileNotFoundException("The input file does not exist.", full);
    }
}
