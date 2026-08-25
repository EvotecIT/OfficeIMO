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
            bool trimTrailingDots = File.Exists(probePath + ".");
            bool trimTrailingSpaces = File.Exists(probePath + " ");
            using (new FileStream(normalizationProbePath, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
            }
            bool normalize = File.Exists(normalizationAlternatePath);
            StringComparer comparer = ignoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
            return normalize || trimTrailingDots || trimTrailingSpaces
                ? new FileNameSemanticsComparer(comparer, normalize, trimTrailingDots, trimTrailingSpaces)
                : comparer;
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

    private static IEqualityComparer<string> DefaultComparer {
        get {
            bool windows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
            StringComparer comparer = windows || RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
                ? StringComparer.OrdinalIgnoreCase
                : StringComparer.Ordinal;
            return windows
                ? new FileNameSemanticsComparer(comparer, normalize: false, trimTrailingDots: true, trimTrailingSpaces: true)
                : comparer;
        }
    }

    private sealed class FileNameSemanticsComparer : IEqualityComparer<string> {
        private readonly StringComparer _comparer;
        private readonly bool _normalize;
        private readonly bool _trimTrailingDots;
        private readonly bool _trimTrailingSpaces;

        internal FileNameSemanticsComparer(
            StringComparer comparer,
            bool normalize,
            bool trimTrailingDots,
            bool trimTrailingSpaces) {
            _comparer = comparer;
            _normalize = normalize;
            _trimTrailingDots = trimTrailingDots;
            _trimTrailingSpaces = trimTrailingSpaces;
        }

        public bool Equals(string? x, string? y) =>
            _comparer.Equals(Canonicalize(x), Canonicalize(y));

        public int GetHashCode(string value) =>
            _comparer.GetHashCode(Canonicalize(value)!);

        private string? Canonicalize(string? value) {
            if (value == null) return null;

            int length = value.Length;
            while (length > 0) {
                char trailing = value[length - 1];
                if ((trailing == '.' && _trimTrailingDots) || (trailing == ' ' && _trimTrailingSpaces)) {
                    length--;
                    continue;
                }
                break;
            }

            string canonical = length == value.Length ? value : value.Substring(0, length);
            return _normalize ? canonical.Normalize(NormalizationForm.FormC) : canonical;
        }
    }
}
