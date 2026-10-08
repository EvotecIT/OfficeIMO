using System;
using System.Collections.Generic;

namespace OfficeIMO.Core.Internal {
    /// <summary>Reusable inert VBA directory and source inspection shared by document-specific storage adapters.</summary>
    internal sealed class OfficeVbaInspection {
        internal OfficeVbaInspection(string limitation) { Limitation = limitation; }
        internal OfficeVbaInspection(string name, int codePage, OfficeVbaModule[] modules, OfficeVbaReference[] references) {
            Name = name; CodePage = codePage; Modules = Array.AsReadOnly(modules); References = Array.AsReadOnly(references);
        }
        internal string? Name { get; }
        internal int? CodePage { get; }
        internal string? Limitation { get; }
        internal IReadOnlyList<OfficeVbaModule> Modules { get; } = Array.Empty<OfficeVbaModule>();
        internal IReadOnlyList<OfficeVbaReference> References { get; } = Array.Empty<OfficeVbaReference>();
    }

    internal sealed class OfficeVbaModule {
        internal OfficeVbaModule(string name, string streamName, bool procedural, int sourceOffset, bool readOnly, bool isPrivate, string? source, string? limitation) {
            Name = name; StreamName = streamName; IsProcedural = procedural; SourceOffset = sourceOffset;
            IsReadOnly = readOnly; IsPrivate = isPrivate; Source = source; Limitation = limitation;
        }
        internal string Name { get; }
        internal string StreamName { get; }
        internal bool IsProcedural { get; }
        internal int SourceOffset { get; }
        internal bool IsReadOnly { get; }
        internal bool IsPrivate { get; }
        internal string? Source { get; }
        internal string? Limitation { get; }
    }

    internal sealed class OfficeVbaReference {
        internal OfficeVbaReference(string name, ushort kind, string libraryId) { Name = name; NativeKind = kind; LibraryId = libraryId; }
        internal string Name { get; }
        internal ushort NativeKind { get; }
        internal string LibraryId { get; }
    }
}
