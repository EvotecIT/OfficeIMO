using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace OfficeIMO.Shared {
    /// <summary>
    /// Validates file-to-file Office conversions without applying format-specific policy.
    /// </summary>
    internal static class OfficeFileConversion {
        internal sealed class Paths {
            internal Paths(string source, string destination) {
                Source = source;
                Destination = destination;
            }

            internal string Source { get; }

            internal string Destination { get; }
        }

        internal static Paths ValidatePaths(
            string sourcePath,
            string destinationPath,
            IReadOnlyCollection<string> supportedExtensions,
            string documentDescription) {
            ValidateExtension(sourcePath, nameof(sourcePath), supportedExtensions, documentDescription);
            ValidateExtension(destinationPath, nameof(destinationPath), supportedExtensions, documentDescription);

            string sourceFullPath = Path.GetFullPath(sourcePath);
            string destinationFullPath = Path.GetFullPath(destinationPath);
            if (!File.Exists(sourceFullPath)) {
                throw new FileNotFoundException($"The source {documentDescription} was not found.", sourceFullPath);
            }

            if (ReferToSameFile(sourceFullPath, destinationFullPath)) {
                throw new IOException("The source and destination paths must be different for conversion.");
            }

            return new Paths(sourceFullPath, destinationFullPath);
        }

        internal static void EnsureDestinationDirectory(string destinationPath) {
            string? directory = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        }

        private static void ValidateExtension(
            string path,
            string parameterName,
            IReadOnlyCollection<string> supportedExtensions,
            string documentDescription) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A file path is required.", parameterName);

            string extension = Path.GetExtension(path);
            if (!supportedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase)) {
                string supported = string.Join(" and ", supportedExtensions);
                throw new NotSupportedException($"{documentDescription} conversion supports {supported} files. The path '{path}' uses '{extension}'.");
            }
        }

        private static bool ReferToSameFile(string left, string right) {
            StringComparison comparison = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;
            return string.Equals(NormalizePath(left), NormalizePath(right), comparison);
        }

        private static string NormalizePath(string path) {
            return path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
    }
}
