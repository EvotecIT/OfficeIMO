using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>Reads reusable, non-executing metadata from an Office VBA compound project.</summary>
    internal static class OfficeVbaProjectInspector {
        private static readonly HashSet<string> InfrastructureStreams = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "dir",
            "_VBA_PROJECT",
            "PROJECT",
            "PROJECTwm"
        };

        internal static IReadOnlyList<string> GetModuleNames(byte[] projectBytes) {
            if (projectBytes == null || projectBytes.Length == 0
                || !OfficeCompoundFileReader.TryRead(projectBytes, out OfficeCompoundFile? compoundFile, out _)
                || compoundFile == null) {
                return Array.Empty<string>();
            }

            return compoundFile.Entries
                .Where(static entry => entry.IsStream && !entry.IsFallback)
                .Where(entry => IsImmediateVbaStream(entry.Path) && !InfrastructureStreams.Contains(entry.Name))
                .Select(static entry => entry.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(static name => name, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        private static bool IsImmediateVbaStream(string path) {
            if (!path.StartsWith("VBA/", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            return path.IndexOf('/', "VBA/".Length) < 0;
        }
    }
}
