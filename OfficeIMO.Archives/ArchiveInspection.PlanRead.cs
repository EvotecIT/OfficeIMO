using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Archives {
    public sealed partial class ArchiveInspection {
        /// <summary>Plans a grouped read from immutable inventory without decoding, checking payload checksums,
        /// consuming quotas or reserving future capacity. Full shared CAB folders count once; separate OpenRead calls
        /// repeat their containing-folder cost. Ordinals must be distinct and readable, as for ReadMembers.
        /// The detached plan survives inspection disposal; actual reads still enforce their current limits.</summary>
        /// <param name="ordinals">Distinct source ordinals in the desired result order.</param>
        /// <param name="cancellationToken">Cancellation while validating and summing the selection.</param>
        /// <exception cref="ArgumentException">The selection repeats an ordinal or exceeds the inventory count.</exception>
        /// <exception cref="ArgumentOutOfRangeException">An ordinal is outside the inventory.</exception>
        /// <exception cref="NotSupportedException">A selected source or compression is unsupported.</exception>
        public ArchiveReadPlan PlanRead(IReadOnlyList<int> ordinals, CancellationToken cancellationToken = default) {
            lock (_gate) {
                ThrowIfDisposed(); cancellationToken.ThrowIfCancellationRequested();
                if (ordinals == null) throw new ArgumentNullException(nameof(ordinals));
                if (ordinals.Count > Entries.Count) throw new ArgumentException("The request exceeds the inventory count.", nameof(ordinals));
                var entries = new ArchiveInspectionEntry[ordinals.Count];
                var seen = new HashSet<int>();
                var folders = new HashSet<(string Cabinet, int Folder)>();
                long selected = 0, decoded = 0, compressed = 0, work = 0, frames = 0;
                for (int index = 0; index < entries.Length; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int ordinal = ordinals[index];
                    if (ordinal < 0 || ordinal >= Entries.Count) throw new ArgumentOutOfRangeException(nameof(ordinals));
                    if (!seen.Add(ordinal)) throw new ArgumentException("A source ordinal is repeated.", nameof(ordinals));
                    var entry = Entries[ordinal];
                    if (entry.Diagnostic != null) throw new NotSupportedException(entry.Diagnostic.Code + ": " + entry.Diagnostic.Message);
                    entries[index] = entry;
                    selected = checked(selected + entry.Length);
                    if (Format == ArchiveFormat.Zip) {
                        decoded = checked(decoded + entry.Length);
                        compressed = checked(compressed + entry.CompressedLength!.Value);
                        work = checked(work + ZipArchiveContent.Work(_zipRecords![ordinal]));
                    } else {
                        string cabinet = entry.CabinetName == null ? string.Empty : entry.CabinetName.Substring(1);
                        if (!folders.Add((cabinet, entry.FolderOrdinal!.Value))) continue;
                        var member = _cabinets[cabinet].Members[entry.CabinetOrdinal!.Value];
                        decoded = checked(decoded + member.FolderExpandedBytes);
                        compressed = checked(compressed + member.FolderCompressedBytes);
                        work = checked(work + member.FolderDecompressionWorkBytes);
                        frames = checked(frames + member.FolderFrameCount);
                    }
                }
                cancellationToken.ThrowIfCancellationRequested();
                return new ArchiveReadPlan(entries, selected, decoded, compressed, work, folders.Count, frames);
            }
        }
    }
}
