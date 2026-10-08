using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Core.Internal {
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

            if (compoundFile.Streams.TryGetValue("VBA/dir", out byte[]? compressed)
                && OfficeVbaCompression.TryDecompress(compressed, 64 * 1024 * 1024, out byte[] directory, out _)
                && OfficeVbaDirectoryCodec.DirectoryModel.TryParse(directory, 64 * 1024 * 1024, out var model, out _, includeSignatureTranscripts: false)
                && model != null) {
                try {
                    return model.Modules.Select(module => module.UnicodeName.Length > 0
                        ? new System.Text.UnicodeEncoding(false, false, true).GetString(module.UnicodeName)
                        : OfficeVbaText.Decode(module.AnsiName, model.CodePage))
                        .OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToArray();
                } catch (System.Text.DecoderFallbackException) { }
                catch (NotSupportedException) { }
            }

            return compoundFile.Entries
                .Where(static entry => entry.IsStream && !entry.IsFallback)
                .Where(entry => IsImmediateVbaStream(entry.Path) && !InfrastructureStreams.Contains(entry.Name)
                    && !entry.Name.StartsWith("__SRP_", StringComparison.OrdinalIgnoreCase))
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
