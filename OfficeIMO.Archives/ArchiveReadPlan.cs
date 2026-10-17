using System.Collections.Generic;

namespace OfficeIMO.Archives {
    /// <summary>Detached declared cost of one grouped member read. Planning neither verifies content nor reserves budgets.</summary>
    public sealed class ArchiveReadPlan {
        internal ArchiveReadPlan(ArchiveInspectionEntry[] entries, long selected, long decoded, long compressed,
            long work, int folders, long frames) {
            Entries = System.Array.AsReadOnly(entries); SelectedBytes = selected; DecodedBytes = decoded;
            CompressedPayloadBytes = compressed; DecompressionWorkBytes = work;
            CabinetFolderCount = folders; CabinetFrameCount = frames;
        }
        /// <summary>Selected immutable inventory records in request order, retaining aliases and duplicate raw names.</summary>
        public IReadOnlyList<ArchiveInspectionEntry> Entries { get; }
        /// <summary>Complete member bytes retained by the grouped read, without coalescing aliases.</summary>
        public long SelectedBytes { get; }
        /// <summary>Declared decoded bytes: selected ZIP members or complete selected CAB folders, each counted once.
        /// Includes unselected folder members and gaps. This is processing volume, not retained decoder memory.</summary>
        public long DecodedBytes { get; }
        /// <summary>Declared compressed payload bytes for selected ZIP members or complete selected CAB folders.
        /// Excludes container headers and does not estimate total I/O, input snapshots or checksum passes.</summary>
        public long CompressedPayloadBytes { get; }
        /// <summary>Work units charged by a successful read under the existing CAB-frame or ZIP-member budget.
        /// Does not predict CPU time. Corruption, cancellation or a quota may stop actual work earlier.</summary>
        public long DecompressionWorkBytes { get; }
        /// <summary>Distinct selected cabinet/folder pairs, including embedded MSI cabinets; zero for ZIP.</summary>
        public int CabinetFolderCount { get; }
        /// <summary>CFDATA frames in those complete folders, each counted once; zero for ZIP.</summary>
        public long CabinetFrameCount { get; }
    }
}
