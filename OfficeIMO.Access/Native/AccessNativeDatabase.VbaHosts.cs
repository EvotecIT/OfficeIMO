using OfficeIMO.Core.Internal;

namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeDatabase {
        internal byte[] GetEventDesigner(AccessApplicationObject host) {
            if (host.StoragePath == null || !_applicationStreams.TryGetValue(host.StoragePath + "Blob", out AccessStorageStream? blob))
                throw new NotSupportedException("The native designer is unavailable.");
            if (_applicationStreams.Any(x => x.Key.StartsWith(host.StoragePath, StringComparison.OrdinalIgnoreCase)
                && (x.Key.EndsWith("BlobDelta", StringComparison.OrdinalIgnoreCase) || x.Key.EndsWith("DeltaBytes", StringComparison.OrdinalIgnoreCase)) && x.Value.Payload.Length != 0))
                throw new NotSupportedException("Designer event changes do not rewrite opaque incremental deltas.");
            return blob.Payload.GetBytes();
        }
        internal bool ApplicationStreamEquals(string path, byte[] bytes) => _applicationStreams.TryGetValue(path, out AccessStorageStream? stream)
            && stream.Payload.GetBytes().SequenceEqual(bytes);

        internal byte[] GetVbaHostProperties(AccessApplicationObject host) {
            if (host.StoragePath == null || !_applicationStreams.TryGetValue(host.StoragePath + "PropData", out AccessStorageStream? stream))
                throw new NotSupportedException("The form/report native properties are unavailable.");
            byte[] metadata = stream.Payload.GetBytes();
            HostModuleValueOffset(metadata);
            return metadata;
        }

        /// <summary>Finds HasModule in the qualified native Long-property records, preserving record order and the optional zero property.</summary>
        internal static int HostModuleValueOffset(byte[] metadata) {
            if (metadata.Length != 13 && metadata.Length != 22 || AccessNativeBinary.I32(metadata, 0) != 0)
                throw new NotSupportedException("The form/report native properties are outside the qualified layout.");
            int found = -1; bool zero = false;
            for (int offset = 4; offset < metadata.Length; offset += 9) {
                if (metadata[offset] != 2) throw new NotSupportedException("The native host property type is unqualified.");
                int code = AccessNativeBinary.I32(metadata, offset + 1), value = AccessNativeBinary.I32(metadata, offset + 5);
                if (code == 1 && found < 0 && (value == 0 || value == 1)) found = offset + 5;
                else if (code == 0 && !zero && value == 0) zero = true;
                else throw new NotSupportedException("The native host property identity or value is unqualified.");
            }
            if (found < 0) throw new NotSupportedException("The native host HasModule property is unavailable.");
            return found;
        }
        /// <summary>Resolves native form/report class bindings independently of generic VBA module kind.</summary>
        internal HashSet<string> GetVbaHostNames(OfficeVbaProject project) {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (AccessApplicationObject host in _document.Forms.Concat(_document.Reports)) {
                string name = (host.CatalogEntry.NativeType == -32768 ? "Form_" : "Report_") + host.Name;
                OfficeVbaModule? module = project.Modules.FirstOrDefault(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                byte[]? metadata = host.StoragePath != null && _applicationStreams.TryGetValue(host.StoragePath + "PropData", out AccessStorageStream? properties)
                    ? properties.Payload.GetBytes() : null;
                int valueOffset;
                try { valueOffset = metadata == null ? -1 : HostModuleValueOffset(metadata); }
                catch (NotSupportedException) { valueOffset = -1; }
                if (valueOffset < 0) {
                    if (module != null) throw new NotSupportedException("A form/report VBA binding has unqualified native host properties.");
                    continue;
                }
                if (AccessNativeBinary.I32(metadata!, valueOffset) == 0) {
                    if (module != null) throw new InvalidDataException("A form/report class has no corresponding native host binding.");
                    continue;
                }
                string? identity = module == null ? null : OfficeVbaText.GetBaseIdentity(module.Source);
                if (module == null || module.Kind != OfficeVbaModuleKind.Document || identity == null || identity.Length != 39
                    || identity[0] != '0' || identity[1] != '{' || identity[38] != '}' || !Guid.TryParse(identity.Substring(1), out _))
                    throw new InvalidDataException("A native form/report binding has no qualified class identity.");
                if (!names.Add(name)) throw new InvalidDataException("Native form/report classes share an ambiguous name.");
            }
            return names;
        }
    }
}
