using OfficeIMO.Core.Internal;

namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeDatabase {
        /// <summary>Resolves native form/report class bindings independently of generic VBA module kind.</summary>
        internal HashSet<string> GetVbaHostNames(OfficeVbaProject project) {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (AccessApplicationObject host in _document.Forms.Concat(_document.Reports)) {
                string name = (host.CatalogEntry.NativeType == -32768 ? "Form_" : "Report_") + host.Name;
                OfficeVbaModule? module = project.Modules.FirstOrDefault(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                byte[]? metadata = host.StoragePath != null && _applicationStreams.TryGetValue(host.StoragePath + "PropData", out AccessStorageStream? properties)
                    ? properties.Payload.GetBytes() : null;
                bool qualified = metadata != null && metadata.Length == 13 && AccessNativeBinary.I32(metadata, 0) == 0
                    && metadata[4] == 2 && AccessNativeBinary.I32(metadata, 5) == 1 && (AccessNativeBinary.I32(metadata, 9) == 0 || AccessNativeBinary.I32(metadata, 9) == 1);
                if (!qualified) {
                    if (module != null) throw new NotSupportedException("A form/report VBA binding has unqualified native host properties.");
                    continue;
                }
                if (AccessNativeBinary.I32(metadata!, 9) == 0) {
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
