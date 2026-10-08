using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Core.Internal;

internal static partial class OfficeVbaProjectInspector {
    /// <summary>Inspects project-relative streams without inventing a compound container or loading compiled code.</summary>
    internal static OfficeVbaInspection Inspect(IReadOnlyDictionary<string, byte[]> streams, int maximumExpandedBytes) {
        if (streams == null) throw new ArgumentNullException(nameof(streams));
        if (maximumExpandedBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumExpandedBytes));
        try { return InspectMetadata(streams, maximumExpandedBytes); }
        catch (Exception exception) when (exception is NotSupportedException || exception is DecoderFallbackException) {
            return new OfficeVbaInspection("The project code page is unavailable or its metadata text is invalid.");
        }
    }

    private static OfficeVbaInspection InspectMetadata(IReadOnlyDictionary<string, byte[]> streams, int maximumExpandedBytes) {
        if (!streams.TryGetValue("VBA/dir", out byte[]? compressed)) return new OfficeVbaInspection("The project has no VBA directory stream.");
        if (!OfficeVbaCompression.TryDecompress(compressed, maximumExpandedBytes, out byte[] directory, out string detail)
            || !OfficeVbaDirectoryCodec.DirectoryModel.TryParse(directory, maximumExpandedBytes, out var model, out detail, includeSignatureTranscripts: false) || model == null)
            return new OfficeVbaInspection(detail);
        int remaining = maximumExpandedBytes - directory.Length;
        var modules = new List<OfficeVbaModuleInspection>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var identities = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var module in model.Modules) {
            string name = module.UnicodeName.Length > 0 ? new UnicodeEncoding(false, false, true).GetString(module.UnicodeName)
                : Decode(module.AnsiName, model.CodePage);
            if (!names.Add(name) || !identities.Add(module.StreamName)) return new OfficeVbaInspection("The VBA directory repeats a module name or stream identity.");
            string? source = null; string? limitation = null;
            if (!streams.TryGetValue("VBA/" + module.StreamName, out byte[]? bytes)) limitation = "The declared module stream is missing.";
            else if (module.TextOffset < 0 || module.TextOffset >= bytes.Length) limitation = "The module has no qualified source container at its recorded offset.";
            else {
                byte[] container = new byte[bytes.Length - module.TextOffset]; Array.Copy(bytes, module.TextOffset, container, 0, container.Length);
                if (!OfficeVbaCompression.TryDecompress(container, remaining, out byte[] expanded, out detail)) limitation = detail;
                else {
                    remaining -= expanded.Length;
                    try { source = Decode(expanded, model.CodePage); }
                    catch (Exception exception) when (exception is NotSupportedException || exception is DecoderFallbackException) {
                        limitation = "The module source code page is unavailable or its text is invalid; the stream remains opaque.";
                    }
                }
            }
            modules.Add(new OfficeVbaModuleInspection(name, module.StreamName, module.TypeId == 0x0021, module.TextOffset,
                module.ReadOnlyRecord != null, module.PrivateRecord != null, source, limitation));
        }
        var references = model.References.Select(reference => new OfficeVbaReferenceInspection(
            reference.UnicodeName.Length > 0 ? new UnicodeEncoding(false, false, true).GetString(reference.UnicodeName) : Decode(reference.AnsiName, model.CodePage),
            reference.Kind, Decode(reference.LibId, model.CodePage))).ToArray();
        return new OfficeVbaInspection(Decode(model.ProjectName, model.CodePage), model.CodePage, modules.ToArray(), references);
    }

    private static string Decode(byte[] bytes, int codePage) => OfficeVbaText.Decode(bytes, codePage).TrimEnd('\0');
}
