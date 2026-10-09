using System.Runtime.CompilerServices;

namespace OfficeIMO.Access {
    public sealed partial class AccessDocument {
        // Weak keys keep detached projects caller-owned while remembering their native
        // identities across intervening applications of other detached projects.
        private readonly ConditionalWeakTable<OfficeVbaModule, VbaModuleIdentity> _vbaModuleIdentities = new ConditionalWeakTable<OfficeVbaModule, VbaModuleIdentity>();
        private sealed class VbaModuleIdentity {
            internal VbaModuleIdentity(Guid id) { Id = id; }
            internal Guid Id { get; }
        }

        private IReadOnlyDictionary<OfficeVbaModule, string> ResolveVbaModuleNames(OfficeVbaProject project) {
            var names = new Dictionary<OfficeVbaModule, string>();
            var current = Catalog.Items.Where(x => x.NativeType == -32761).ToDictionary(x => x.Id);
            foreach (OfficeVbaModule module in project.Modules) {
                if (!_vbaModuleIdentities.TryGetValue(module, out VbaModuleIdentity? binding)) continue;
                if (!current.TryGetValue(binding.Id, out AccessCatalogEntry? entry))
                    throw new InvalidOperationException("A detached VBA module identity was removed from this database. Reload its current project before editing.");
                names.Add(module, entry.Name);
            }
            return names;
        }

        private Action? BindVbaModuleIdentities(OfficeVbaProject project, IEnumerable<AccessCatalogEntry> catalog, bool undoable) {
            var ordinary = catalog.Where(x => x.NativeType == -32761).ToDictionary(x => x.Name, StringComparer.OrdinalIgnoreCase);
            var previous = undoable ? new List<KeyValuePair<OfficeVbaModule, VbaModuleIdentity?>>() : null;
            foreach (OfficeVbaModule module in project.Modules) {
                if (module.Kind != OfficeVbaModuleKind.Standard && module.Kind != OfficeVbaModuleKind.Class
                    || !ordinary.TryGetValue(module.Name, out AccessCatalogEntry? entry)) continue;
                _vbaModuleIdentities.TryGetValue(module, out VbaModuleIdentity? before);
                previous?.Add(new KeyValuePair<OfficeVbaModule, VbaModuleIdentity?>(module, before));
                _vbaModuleIdentities.Remove(module);
                _vbaModuleIdentities.Add(module, new VbaModuleIdentity(entry.Id));
            }
            if (previous == null) return null;
            return () => {
                foreach (var binding in previous) {
                    _vbaModuleIdentities.Remove(binding.Key);
                    if (binding.Value != null) _vbaModuleIdentities.Add(binding.Key, binding.Value);
                }
            };
        }
    }
}
