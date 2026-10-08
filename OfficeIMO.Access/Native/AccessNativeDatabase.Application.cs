using OfficeIMO.Core.Internal;
using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeDatabase {
        private void LoadApplicationObjects(CancellationToken cancellation) {
            Dictionary<string, AccessStorageStream> streams = ReadApplicationStorage(cancellation);
            LoadApplicationCollection(_document.Forms, -32768, "Forms", streams, cancellation);
            LoadApplicationCollection(_document.Reports, -32764, "Reports", streams, cancellation);
            LoadApplicationCollection(_document.Macros, -32766, "Scripts", streams, cancellation);
            LoadDataMacros(cancellation);
            LoadResources(cancellation);
            LoadDependencies();
            const string prefix = "VBA/VBAProject/";
            Dictionary<string, byte[]> projectStreams = streams.Where(x => x.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                .ToDictionary(x => x.Key.Substring(prefix.Length), x => x.Value.Payload.GetBytes(), StringComparer.OrdinalIgnoreCase);
            if (projectStreams.Count == 0) {
                _document.VbaProject = new AccessVbaProjectInfo(_catalog.Any(x => x.Type == -32761) ? AccessCatalogStatus.NotDecoded : AccessCatalogStatus.Decoded);
                return;
            }
            OfficeVbaInspection inspection;
            try { inspection = OfficeVbaProjectInspector.Inspect(projectStreams, MaxMetadataBytes, cancellation); }
            catch (Exception exception) when (exception is NotSupportedException || exception is System.Text.DecoderFallbackException) {
                inspection = new OfficeVbaInspection("The project directory contains an unsupported text encoding; its native streams remain preserve-only.");
            }
            _document.VbaProject = new AccessVbaProjectInfo(inspection.Limitation == null ? AccessCatalogStatus.Decoded : AccessCatalogStatus.NotDecoded) {
                Name = inspection.Name, CodePage = inspection.CodePage,
                Modules = Array.AsReadOnly(inspection.Modules.Select(x => new AccessVbaModuleInfo(x, prefix)).ToArray()),
                References = Array.AsReadOnly(inspection.References.Select(x => new AccessVbaReferenceInfo(x)).ToArray()),
                Diagnostics = inspection.Limitation == null ? Array.AsReadOnly(Array.Empty<AccessDiagnostic>())
                    : Array.AsReadOnly(new[] { new AccessDiagnostic("access.vba.directory-opaque", inspection.Limitation) })
            };
        }
        private void LoadApplicationCollection(AccessObjectCollection<AccessApplicationObject> collection, int catalogType, string group,
            Dictionary<string, AccessStorageStream> streams, CancellationToken cancellation) {
            Dictionary<string, int>? slots = null;
            if (streams.TryGetValue(group + "/\u0003DirData", out AccessStorageStream? directory)) {
                AccountMetadata(directory.Payload.Length);
                slots = ReadObjectDirectory(directory.Payload.GetBytes());
            }
            foreach (AccessCatalogEntry? catalog in _document.Catalog.Where(x => x.NativeType == catalogType)) {
                cancellation.ThrowIfCancellationRequested();
                string? path = slots != null && slots.TryGetValue(catalog.Name, out int slot) ? group + "/" + slot.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/" : null;
                AccessApplicationObject model = new AccessApplicationObject(_document, catalog.Name, catalog, path) {
                    Streams = path == null ? Array.AsReadOnly(Array.Empty<AccessStorageStream>())
                        : Array.AsReadOnly(streams.Where(x => x.Key.StartsWith(path, StringComparison.OrdinalIgnoreCase)).Select(x => x.Value).ToArray())
                };
                if (path != null && streams.TryGetValue(path + "Blob", out AccessStorageStream? blob)) {
                    AccountMetadata(blob.Payload.Length);
                    byte[] payload = blob.Payload.GetBytes();
                    if (catalogType == -32766) model.ActionMacro = AccessNativeActionMacro.Read(payload);
                    else model.Definition = AccessNativeDesigner.Read(payload, MaxCatalogObjects, cancellation);
                }
                model.Diagnostics = Array.AsReadOnly(new[] { new AccessDiagnostic(model.ActionMacro != null ? "access.application.action-macro-inert" : model.Definition == null ? "access.application.preserve-only" : "access.application.partial-designer",
                    model.ActionMacro != null ? "The qualified single StopMacro action is read without execution; other action layouts remain preserve-only."
                        : model.Definition == null ? "This application's native payload remains preserve-only; typed designer semantics are unqualified."
                        : "The expanded designer tree is read inertly. Unknown/default properties and incremental deltas remain preserve-only; no events, expressions or controls are instantiated.", model.Id) });
                collection.AddNativeItem(model);
            }
            collection.CatalogStatus = AccessCatalogStatus.Decoded;
        }
        private static Dictionary<string, int>? ReadObjectDirectory(byte[] bytes) {
            Dictionary<string, int> result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            HashSet<int> slots = new HashSet<int>();
            if (bytes.Length < 4 || U32(bytes, 0) != 0) return null;
            int offset = 4;
            while (offset < bytes.Length) {
                if (bytes.Length - offset < 2 || bytes[offset] != 4) return null;
                int length = bytes[offset + 1]; offset += 2;
                if (length < 6 || (length & 1) != 0 || offset > bytes.Length - length) return null;
                string name;
                try { name = new System.Text.UnicodeEncoding(false, false, true).GetString(bytes, offset, length - 4); }
                catch (System.Text.DecoderFallbackException) { return null; }
                int slot = I32(bytes, offset + length - 4); offset += length;
                if (slot < 0 || !slots.Add(slot)) return null;
                if (result.ContainsKey(name)) return null; result.Add(name, slot);
            }
            return result;
        }
    }
}
