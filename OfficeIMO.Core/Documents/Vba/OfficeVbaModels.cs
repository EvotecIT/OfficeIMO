using System;
using System.Collections.Generic;

namespace OfficeIMO;

/// <summary>The identity and persistence role of a VBA source module.</summary>
public enum OfficeVbaModuleKind {
    /// <summary>A procedural module exported as .bas.</summary>
    Standard,
    /// <summary>An ordinary class module exported as .cls.</summary>
    Class,
    /// <summary>A host-owned document, worksheet, or workbook class.</summary>
    Document,
    /// <summary>A class coupled to an existing form designer; source edits preserve its opaque design.</summary>
    Designer
}

/// <summary>A detached source module belonging to one editable VBA project.</summary>
public sealed class OfficeVbaModule {
    internal OfficeVbaModule(string name, OfficeVbaModuleKind kind, string source,
        Core.Internal.OfficeVbaDirectoryCodec.ModuleModel directory) {
        Name = name; OriginalName = name; Kind = kind; Source = source; OriginalSource = source; Directory = directory;
    }
    /// <summary>Gets the logical module name, independent of its compound stream name.</summary>
    public string Name { get; internal set; }
    /// <summary>Gets the module's persistence role.</summary>
    public OfficeVbaModuleKind Kind { get; }
    /// <summary>Gets the complete source, including the persisted Attribute lines.</summary>
    public string Source { get; internal set; }
    internal string OriginalName { get; }
    internal string OriginalSource { get; }
    internal Core.Internal.OfficeVbaDirectoryCodec.ModuleModel Directory { get; }
    internal bool IsNew { get; set; }
    internal bool IsDocumentClass { get; set; }
}

/// <summary>A VBA library reference; unknown/control/project reference records remain opaque and preserved.</summary>
public sealed class OfficeVbaReference {
    internal OfficeVbaReference(string name, string? libraryId, byte[] serialized) {
        Name = name; LibraryId = libraryId; Serialized = serialized;
    }
    /// <summary>Gets the reference name when the file supplies one.</summary>
    public string Name { get; }
    /// <summary>Gets a registered type-library identifier, or null for another reference kind.</summary>
    public string? LibraryId { get; }
    internal byte[] Serialized { get; }
}

/// <summary>Limits the materialized compound streams and expanded module source.</summary>
public sealed class OfficeVbaReadOptions {
    /// <summary>Gets or sets the maximum input and aggregate compound-stream bytes; defaults to 64 MiB.</summary>
    public int MaximumProjectBytes { get; set; } = 64 * 1024 * 1024;
    /// <summary>Gets or sets the aggregate expanded directory and source limit; defaults to 64 MiB.</summary>
    public int MaximumExpandedBytes { get; set; } = 64 * 1024 * 1024;
}

/// <summary>Controls explicit signature invalidation and bounded VBA output.</summary>
public sealed class OfficeVbaWriteOptions {
    /// <summary>Gets or sets whether a document adapter may remove its existing VBA signatures; defaults to rejection.</summary>
    public bool AllowSignatureRemoval { get; set; }
    /// <summary>Gets or sets the maximum complete compound-file output size; defaults to 64 MiB.</summary>
    public int MaximumProjectBytes { get; set; } = 64 * 1024 * 1024;
    /// <summary>Gets or sets the aggregate existing VBA payload and child-part bytes retained for document-adapter recovery; defaults to 64 MiB.</summary>
    public int MaximumRecoveryBytes { get; set; } = 64 * 1024 * 1024;
    /// <summary>Gets or sets the aggregate expanded directory and source limit; defaults to 64 MiB.</summary>
    public int MaximumExpandedBytes { get; set; } = 64 * 1024 * 1024;
}

/// <summary>Reports the compound artifact and the consequences of a VBA write.</summary>
public sealed class OfficeVbaWriteResult {
    internal OfficeVbaWriteResult(byte[] bytes, bool changed, IReadOnlyList<string> changedModules) {
        _bytes = bytes; Changed = changed; ChangedModules = changedModules;
    }
    private readonly byte[] _bytes;
    /// <summary>Gets a caller-owned copy of the complete vbaProject.bin bytes.</summary>
    public byte[] GetBytes() => (byte[])_bytes.Clone();
    /// <summary>Gets whether the result changes the loaded artifact.</summary>
    public bool Changed { get; }
    /// <summary>Gets the edited, added, renamed, deleted, and compression-normalized module names.</summary>
    public IReadOnlyList<string> ChangedModules { get; }
}
