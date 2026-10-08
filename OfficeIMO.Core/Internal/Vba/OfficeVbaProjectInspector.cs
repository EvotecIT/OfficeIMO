using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Security;

namespace OfficeIMO.Core.Internal {
    /// <summary>Reads reusable, non-executing metadata from an Office VBA compound project.</summary>
    internal static class OfficeVbaProjectInspector {
        /// <summary>Inspects project-relative streams supplied by a document storage adapter. No compiled code is loaded.</summary>
        internal static OfficeVbaInspection Inspect(IReadOnlyDictionary<string, byte[]> streams, int maximumExpandedBytes) {
            if (streams == null) throw new ArgumentNullException(nameof(streams));
            if (maximumExpandedBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumExpandedBytes));
            if (!streams.TryGetValue("VBA/dir", out byte[]? compressed)) return new OfficeVbaInspection("The project has no VBA directory stream.");
            if (!OfficeVbaProjectCanonicalizer.TryDecompress(compressed, maximumExpandedBytes, out byte[] directory, out string detail)
                || !OfficeVbaProjectCanonicalizer.DirectoryModel.TryParse(directory, maximumExpandedBytes, out var model, out detail) || model == null)
                return new OfficeVbaInspection(detail);
            int remaining = maximumExpandedBytes - directory.Length;
            var modules = new List<OfficeVbaModule>();
            foreach (var module in model.Modules) {
                string name = module.UnicodeName.Length > 0 ? new UnicodeEncoding(false, false, true).GetString(module.UnicodeName)
                    : Decode(module.AnsiName, model.CodePage);
                string? source = null; string? limitation = null;
                if (!streams.TryGetValue("VBA/" + module.StreamName, out byte[]? bytes)) limitation = "The declared module stream is missing.";
                else if (module.TextOffset < 0 || module.TextOffset >= bytes.Length) limitation = "The module has no qualified source container at its recorded offset.";
                else {
                    byte[] container = new byte[bytes.Length - module.TextOffset]; Array.Copy(bytes, module.TextOffset, container, 0, container.Length);
                    if (!OfficeVbaProjectCanonicalizer.TryDecompress(container, remaining, out byte[] expanded, out detail)) limitation = detail;
                    else {
                        remaining -= expanded.Length;
                        try { source = Decode(expanded, model.CodePage); }
                        catch (NotSupportedException) { limitation = "The module source code page is not supported; its stream remains opaque."; }
                    }
                }
                modules.Add(new OfficeVbaModule(name, module.StreamName, module.TypeId == 0x0021, module.TextOffset,
                    module.ReadOnlyRecord != null, module.PrivateRecord != null, source, limitation));
            }
            var references = model.References.Select(reference => new OfficeVbaReference(
                new UnicodeEncoding(false, false, true).GetString(reference.UnicodeName), reference.Kind,
                Decode(reference.LibId, model.CodePage))).ToArray();
            return new OfficeVbaInspection(Decode(model.ProjectName, model.CodePage), model.CodePage, modules.ToArray(), references);
        }

        private static string Decode(byte[] bytes, int codePage) => bytes.All(x => x < 128) ? Encoding.ASCII.GetString(bytes).TrimEnd('\0')
            : codePage == 65001 ? new UTF8Encoding(false, true).GetString(bytes)
            : OfficeLegacySingleByteEncoding.Decode(bytes, 0, bytes.Length, codePage).TrimEnd('\0');
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
