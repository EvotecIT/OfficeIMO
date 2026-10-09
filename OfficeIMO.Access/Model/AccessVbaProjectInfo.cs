namespace OfficeIMO.Access {
    /// <summary>Inert VBA project inspection. Source extraction does not compile code, validate signatures or recover missing compiled source.</summary>
    public sealed class AccessVbaProjectInfo {
        internal AccessVbaProjectInfo(AccessCatalogStatus status) { CatalogStatus = status; }
        /// <summary>Availability of the qualified project directory inventory.</summary>
        public AccessCatalogStatus CatalogStatus { get; }
        /// <summary>Project name recorded in the VBA directory.</summary>
        public string? Name { get; internal set; }
        /// <summary>Code page used by the project source and ANSI directory records.</summary>
        public int? CodePage { get; internal set; }
        /// <summary>Declared modules, including entries without recoverable source.</summary>
        public IReadOnlyList<AccessVbaModuleInfo> Modules { get; internal set; } = Array.AsReadOnly(Array.Empty<AccessVbaModuleInfo>());
        /// <summary>Module names from the declared directory, rather than inferred opaque stream names.</summary>
        public IReadOnlyList<string> ModuleNames => Array.AsReadOnly(Modules.Select(x => x.Name).ToArray());
        /// <summary>Persisted library references; referenced libraries are never opened or loaded.</summary>
        public IReadOnlyList<AccessVbaReferenceInfo> References { get; internal set; } = Array.AsReadOnly(Array.Empty<AccessVbaReferenceInfo>());
        /// <summary>Inspection limitations, including unknown source containers or code pages.</summary>
        public IReadOnlyList<AccessDiagnostic> Diagnostics { get; internal set; } = Array.AsReadOnly(Array.Empty<AccessDiagnostic>());
    }

    /// <summary>A declared VBA module and its available inert source.</summary>
    public sealed class AccessVbaModuleInfo {
        internal AccessVbaModuleInfo(OfficeIMO.Core.Internal.OfficeVbaModuleInspection module, string storagePrefix) {
            Name = module.Name; StoragePath = storagePrefix + "VBA/" + module.StreamName; IsProcedural = module.IsProcedural;
            SourceOffset = module.SourceOffset; IsReadOnly = module.IsReadOnly; IsPrivate = module.IsPrivate;
            Source = module.Source; SourceLimitation = module.Limitation;
        }
        /// <summary>Declared module name.</summary>
        public string Name { get; }
        /// <summary>Native storage stream path, separate from the module name.</summary>
        public string StoragePath { get; }
        /// <summary>Whether the directory marks a standard procedural module.</summary>
        public bool IsProcedural { get; }
        /// <summary>Recorded source-container offset after compiled/cache content.</summary>
        public int SourceOffset { get; }
        /// <summary>Persisted read-only module flag.</summary>
        public bool IsReadOnly { get; }
        /// <summary>Persisted private module flag.</summary>
        public bool IsPrivate { get; }
        /// <summary>Available source including attributes; null means it cannot be recovered by the qualified decoder.</summary>
        public string? Source { get; }
        /// <summary>Reason source is unavailable; contains no source text.</summary>
        public string? SourceLimitation { get; }
    }

    /// <summary>An inert VBA reference with its native identifier.</summary>
    public sealed class AccessVbaReferenceInfo {
        internal AccessVbaReferenceInfo(OfficeIMO.Core.Internal.OfficeVbaReferenceInspection reference) { Name = reference.Name; NativeKind = reference.NativeKind; LibraryId = reference.LibraryId; }
        /// <summary>Declared reference name.</summary>
        public string Name { get; }
        /// <summary>Native reference record type.</summary>
        public ushort NativeKind { get; }
        /// <summary>Stored library identifier or path. This is explicit native metadata and is never resolved.</summary>
        public string LibraryId { get; }
    }
}
