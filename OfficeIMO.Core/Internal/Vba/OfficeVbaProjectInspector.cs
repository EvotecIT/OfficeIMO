using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Core.Internal {
    /// <summary>Reads reusable, non-executing metadata from an Office VBA compound project.</summary>
    internal static class OfficeVbaProjectInspector {
        /// <summary>Inspects project-relative streams without executing code. Encoded processing is bounded by the supplied stream bytes, separately from the expansion limit.</summary>
        internal static OfficeVbaInspection Inspect(IReadOnlyDictionary<string, byte[]> streams, int maximumExpandedBytes, CancellationToken cancellationToken = default) {
            if (streams == null) throw new ArgumentNullException(nameof(streams));
            if (maximumExpandedBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumExpandedBytes));
            cancellationToken.ThrowIfCancellationRequested();
            int remaining = maximumExpandedBytes;
            if (!streams.TryGetValue("VBA/dir", out byte[]? compressed)) return new OfficeVbaInspection("The project has no VBA directory stream.");
            long encodedRemaining = -compressed.Length;
            foreach (byte[] stream in streams.Values) {
                cancellationToken.ThrowIfCancellationRequested();
                encodedRemaining += stream.Length;
            }
            if (!OfficeVbaCompression.TryDecompress(compressed, ref remaining, out byte[] directory, out string detail, cancellationToken)
                || !OfficeVbaDirectoryCodec.DirectoryModel.TryParse(directory, maximumExpandedBytes, out OfficeVbaDirectoryCodec.DirectoryModel? model, out detail, includeSignatureTranscripts: false) || model == null)
                return new OfficeVbaInspection(detail);
            List<OfficeVbaModuleInspection> modules = new List<OfficeVbaModuleInspection>();
            foreach (OfficeVbaDirectoryCodec.ModuleModel module in model.Modules) {
                cancellationToken.ThrowIfCancellationRequested();
                string name = module.UnicodeName.Length > 0 ? new UnicodeEncoding(false, false, true).GetString(module.UnicodeName)
                    : Decode(module.AnsiName, model.CodePage);
                string? source = null; string? limitation = null;
                if (!streams.TryGetValue("VBA/" + module.StreamName, out byte[]? bytes)) limitation = "The declared module stream is missing.";
                else if (module.TextOffset < 0 || module.TextOffset >= bytes.Length) limitation = "The module has no qualified source container at its recorded offset.";
                else if (remaining == 0) limitation = "The expanded MS-OVBA project exceeds the configured byte limit; this module remains opaque.";
                else if (bytes.Length - module.TextOffset > encodedRemaining) limitation = "The MS-OVBA project exceeds its encoded input byte limit; this module remains opaque.";
                else {
                    // Charge each attempt before copying, including immediate failures and repeated declarations.
                    encodedRemaining -= bytes.Length - module.TextOffset;
                    byte[] container = new byte[bytes.Length - module.TextOffset]; Array.Copy(bytes, module.TextOffset, container, 0, container.Length);
                    if (!OfficeVbaCompression.TryDecompress(container, ref remaining, out byte[] expanded, out detail, cancellationToken)) limitation = detail;
                    else {
                        try { source = Decode(expanded, model.CodePage); }
                        catch (NotSupportedException) { limitation = "The module source code page is not supported; its stream remains opaque."; }
                        catch (DecoderFallbackException) { limitation = "The module source is not valid for its declared code page; its stream remains opaque."; }
                    }
                }
                modules.Add(new OfficeVbaModuleInspection(name, module.StreamName, module.TypeId == 0x0021, module.TextOffset,
                    module.ReadOnlyRecord != null, module.PrivateRecord != null, source, limitation));
            }
            cancellationToken.ThrowIfCancellationRequested();
            OfficeVbaReferenceInspection[] references = model.References.Select(reference => new OfficeVbaReferenceInspection(
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

            if (TryReadDirectory(compoundFile, out OfficeVbaDirectoryCodec.DirectoryModel? model)) {
                try {
                    return model!.Modules.Select(module => module.UnicodeName.Length > 0
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

        /// <summary>Checks directory structure before choosing metadata-based or opaque stream operations.</summary>
        internal static bool TryReadDirectory(OfficeCompoundFile compoundFile, out OfficeVbaDirectoryCodec.DirectoryModel? model) {
            model = null;
            return compoundFile.Streams.TryGetValue("VBA/dir", out byte[]? compressed)
                && OfficeVbaCompression.TryDecompress(compressed, 64 * 1024 * 1024, out byte[] directory, out _)
                && OfficeVbaDirectoryCodec.DirectoryModel.TryParse(directory, 64 * 1024 * 1024, out model, out _, includeSignatureTranscripts: false)
                && model != null;
        }

        private static bool IsImmediateVbaStream(string path) {
            if (!path.StartsWith("VBA/", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            return path.IndexOf('/', "VBA/".Length) < 0;
        }
    }
}
