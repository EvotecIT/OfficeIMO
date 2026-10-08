using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Core.Internal {

    internal static partial class OfficeVbaProjectInspector {
        /// <summary>Inspects project-relative streams without inventing a compound container or loading compiled code.</summary>
        internal static OfficeVbaInspection Inspect(IReadOnlyDictionary<string, byte[]> streams, int maximumExpandedBytes,
            CancellationToken cancellationToken = default) {
            if (streams == null) throw new ArgumentNullException(nameof(streams));
            if (maximumExpandedBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumExpandedBytes));
            cancellationToken.ThrowIfCancellationRequested();
            try { return InspectMetadata(streams, maximumExpandedBytes, cancellationToken); }
            catch (Exception exception) when (exception is NotSupportedException || exception is DecoderFallbackException) {
                return new OfficeVbaInspection("The project code page is unavailable or its metadata text is invalid.");
            }
        }

        private static OfficeVbaInspection InspectMetadata(IReadOnlyDictionary<string, byte[]> streams, int maximumExpandedBytes,
            CancellationToken cancellationToken) {
            if (!streams.TryGetValue("VBA/dir", out byte[]? compressed)) return new OfficeVbaInspection("The project has no VBA directory stream.");
            int remaining = maximumExpandedBytes;
            if (!OfficeVbaCompression.TryDecompress(compressed, ref remaining, out byte[] directory, out string detail, cancellationToken)
                || !OfficeVbaDirectoryCodec.DirectoryModel.TryParse(directory, maximumExpandedBytes, out OfficeVbaDirectoryCodec.DirectoryModel? model, out detail, includeSignatureTranscripts: false) || model == null)
                return new OfficeVbaInspection(detail);
            List<OfficeVbaModuleInspection> modules = new List<OfficeVbaModuleInspection>();
            HashSet<string> names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> identities = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (OfficeVbaDirectoryCodec.ModuleModel module in model.Modules) {
                cancellationToken.ThrowIfCancellationRequested();
                string name = module.UnicodeName.Length > 0 ? new UnicodeEncoding(false, false, true).GetString(module.UnicodeName)
                    : Decode(module.AnsiName, model.CodePage);
                if (!names.Add(name) || !identities.Add(module.StreamName)) return new OfficeVbaInspection("The VBA directory repeats a module name or stream identity.");
                string? source = null; string? limitation = null;
                if (!streams.TryGetValue("VBA/" + module.StreamName, out byte[]? bytes)) limitation = "The declared module stream is missing.";
                else if (module.TextOffset < 0 || module.TextOffset >= bytes.Length) limitation = "The module has no qualified source container at its recorded offset.";
                else if (remaining == 0) limitation = "The expanded MS-OVBA project exceeds the configured byte limit; this module remains opaque.";
                else {
                    byte[] container = new byte[bytes.Length - module.TextOffset]; Array.Copy(bytes, module.TextOffset, container, 0, container.Length);
                    if (!OfficeVbaCompression.TryDecompress(container, ref remaining, out byte[] expanded, out detail, cancellationToken)) limitation = detail;
                    else {
                        try { source = Decode(expanded, model.CodePage); }
                        catch (Exception exception) when (exception is NotSupportedException || exception is DecoderFallbackException) {
                            limitation = "The module source code page is unavailable or its text is invalid; the stream remains opaque.";
                        }
                    }
                }
                modules.Add(new OfficeVbaModuleInspection(name, module.StreamName, module.TypeId == 0x0021, module.TextOffset,
                    module.ReadOnlyRecord != null, module.PrivateRecord != null, source, limitation));
            }
            cancellationToken.ThrowIfCancellationRequested();
            OfficeVbaReferenceInspection[] references = model.References.Select(reference => new OfficeVbaReferenceInspection(
                reference.UnicodeName.Length > 0 ? new UnicodeEncoding(false, false, true).GetString(reference.UnicodeName) : Decode(reference.AnsiName, model.CodePage),
                reference.Kind, Decode(reference.LibId, model.CodePage))).ToArray();
            return new OfficeVbaInspection(Decode(model.ProjectName, model.CodePage), model.CodePage, modules.ToArray(), references);
        }

        private static string Decode(byte[] bytes, int codePage) => OfficeVbaText.Decode(bytes, codePage).TrimEnd('\0');
    }
}
