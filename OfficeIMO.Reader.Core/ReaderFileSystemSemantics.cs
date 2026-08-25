using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeIMO.Reader;

/// <summary>
/// Resolves destination-directory filename semantics for collision-safe materialization.
/// </summary>
internal static class ReaderFileSystemSemantics {
    internal static IEqualityComparer<string> GetFileNameComparer(string directoryPath) {
        string probeName = ".officeimo-case-probe-" + Guid.NewGuid().ToString("N") + "a";
        string alternateName = probeName.Substring(0, probeName.Length - 1) + "A";
        string probePath = Path.Combine(directoryPath, probeName);
        string alternatePath = Path.Combine(directoryPath, alternateName);
        string normalizationProbeName = ".officeimo-normalization-probe-" + Guid.NewGuid().ToString("N") + "-\u00e9";
        string normalizationAlternateName = normalizationProbeName.Normalize(NormalizationForm.FormD);
        string normalizationProbePath = Path.Combine(directoryPath, normalizationProbeName);
        string normalizationAlternatePath = Path.Combine(directoryPath, normalizationAlternateName);
        try {
            using (new FileStream(probePath, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
            }
            bool ignoreCase = File.Exists(alternatePath);
            using (new FileStream(normalizationProbePath, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
            }
            bool normalize = File.Exists(normalizationAlternatePath);
            StringComparer comparer = ignoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
            return normalize ? new NormalizedFileNameComparer(comparer) : comparer;
        } catch (IOException) {
            return DefaultComparer;
        } catch (UnauthorizedAccessException) {
            return DefaultComparer;
        } finally {
            try {
                if (File.Exists(probePath)) File.Delete(probePath);
            } catch {
                // Best-effort cleanup must not hide the caller's materialization result.
            }
            try {
                if (File.Exists(normalizationProbePath)) File.Delete(normalizationProbePath);
            } catch {
                // Best-effort cleanup must not hide the caller's materialization result.
            }
        }
    }

    private static StringComparer DefaultComparer =>
        RuntimeInformation.IsOSPlatform(OSPlatform.Windows) || RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
            ? StringComparer.OrdinalIgnoreCase
            : StringComparer.Ordinal;

    private sealed class NormalizedFileNameComparer : IEqualityComparer<string> {
        private readonly StringComparer _comparer;

        internal NormalizedFileNameComparer(StringComparer comparer) {
            _comparer = comparer;
        }

        public bool Equals(string? x, string? y) =>
            _comparer.Equals(Normalize(x), Normalize(y));

        public int GetHashCode(string value) =>
            _comparer.GetHashCode(Normalize(value)!);

        private static string? Normalize(string? value) => value?.Normalize(NormalizationForm.FormC);
    }
}
