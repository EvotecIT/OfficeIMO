using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Archives;

/// <summary>Complete selected members from one atomic grouped read. Dispose closes its streams and releases its member collection.</summary>
public sealed class ArchiveReadBatch : IDisposable {
    private readonly List<ArchiveReadMember> _members;
    internal ArchiveReadBatch(List<ArchiveReadMember> members, long selectedBytes, long decodedFolderBytes, long workBytes) {
        _members = members; Members = members.AsReadOnly();
        SelectedBytes = selectedBytes; DecodedFolderBytes = decodedFolderBytes; DecompressionWorkBytes = workBytes;
    }
    /// <summary>Complete members in the requested ordinal order. Empty after disposal.</summary>
    public IReadOnlyList<ArchiveReadMember> Members { get; }
    /// <summary>Complete selected payload bytes retained by this batch, without coalescing duplicate names or aliases.</summary>
    public long SelectedBytes { get; }
    /// <summary>Bytes decoded across the selected folders, each counted once, including unselected members and folder gaps.</summary>
    public long DecodedFolderBytes { get; }
    /// <summary>CAB frame or ZIP processing work charged by this read, using the inspection's bounded work units rather than exact CPU time.</summary>
    public long DecompressionWorkBytes { get; }
    /// <summary>Closes every member stream. Inspection disposal does not close these streams; this batch owns their lifetime.</summary>
    public void Dispose() {
        foreach (ArchiveReadMember member in _members) member.Content.Dispose();
        _members.Clear();
    }
}

/// <summary>One selected authored record with its complete read-only content. Its containing batch owns the stream.</summary>
public sealed class ArchiveReadMember {
    internal ArchiveReadMember(ArchiveInspectionEntry entry, Stream content) { Entry = entry; Content = content; }
    /// <summary>The original immutable inventory record, retaining raw identity and source ordinal.</summary>
    public ArchiveInspectionEntry Entry { get; }
    /// <summary>Complete length-validated member bytes, initially positioned at zero. Independently usable after inspection disposal.</summary>
    public Stream Content { get; }
}
