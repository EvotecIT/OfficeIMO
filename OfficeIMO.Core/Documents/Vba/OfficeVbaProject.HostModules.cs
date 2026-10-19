using System;

namespace OfficeIMO;

public sealed partial class OfficeVbaProject {
    /// <summary>Removes a document module after its host has removed the corresponding document object.</summary>
    internal void DeleteDocumentModule(string name) {
        EnsureEditable();
        OfficeVbaModule module = GetModule(name);
        if (module.Kind != OfficeVbaModuleKind.Document)
            throw new ArgumentException("The host can remove only a document module through this operation.", nameof(name));
        if (!module.IsNew) _deletedModules.Add(module.OriginalName);
        _modules.Remove(module);
    }
}
