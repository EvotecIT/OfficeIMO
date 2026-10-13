using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using OfficeIMO.Core.Internal.Cabinet;

namespace OfficeIMO.Archives;

public sealed partial class ArchiveInspection {
    /// <summary>Reads selected complete members in request order, decoding each selected cabinet folder once.
    /// Raw duplicate names and MSI aliases remain separate. Selected payload size, ordinals and source support are checked
    /// before decoding. A corruption, limit or cancellation failure returns no batch or partial member.
    /// The inspection's cumulative expansion and frame-processing work limits also apply. A batch owns its read-only streams
    /// independently of this inspection. Memory includes retained selected bytes, input and decoder buffers; limits are not a peak-RSS guarantee.</summary>
    /// <param name="ordinals">Distinct source ordinals from Entries. Their order determines the returned member order.</param>
    /// <param name="maxSelectedBytes">Finite maximum aggregate selected payload bytes retained in this batch; zero permits only empty members.</param>
    /// <param name="cancellationToken">Cancellation for preflight, folder decoding and selected payload copies.</param>
    /// <exception cref="OfficeInspectionResourceLimitException">The selected payload or remaining expansion/work allowance is exhausted.</exception>
    /// <exception cref="ArgumentException">An ordinal is repeated or the request exceeds the inventory count.</exception>
    /// <exception cref="ArgumentOutOfRangeException">A size limit is negative or an ordinal is outside the inventory.</exception>
    /// <exception cref="NotSupportedException">A selected source or compression cannot be read.</exception>
    /// <exception cref="InvalidDataException">Selected payloads exceed their limit, decoded content is corrupt or a cumulative inspection limit is exceeded.</exception>
    public ArchiveReadBatch ReadMembers(IReadOnlyList<int> ordinals, long maxSelectedBytes,
        CancellationToken cancellationToken = default) {
        lock (_gate) {
            ThrowIfDisposed(); cancellationToken.ThrowIfCancellationRequested();
            if (ordinals == null) throw new ArgumentNullException(nameof(ordinals));
            if (maxSelectedBytes < 0) throw new ArgumentOutOfRangeException(nameof(maxSelectedBytes));
            int count = ordinals.Count;
            if (count > Entries.Count) throw new ArgumentException("The request exceeds the inventory count.", nameof(ordinals));
            var entries = new ArchiveInspectionEntry[count];
            var seen = new HashSet<int>();
            var groups = new Dictionary<(string Cabinet, int Folder), SelectedFolder>();
            var orderedGroups = new List<SelectedFolder>();
            long selectedBytes = 0;
            long expandedBytes = 0;
            for (int index = 0; index < count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                int ordinal = ordinals[index];
                if (ordinal < 0 || ordinal >= Entries.Count) throw new ArgumentOutOfRangeException(nameof(ordinals));
                if (!seen.Add(ordinal)) throw new ArgumentException("A source ordinal is repeated.", nameof(ordinals));
                ArchiveInspectionEntry entry = Entries[ordinal];
                if (entry.Diagnostic != null) throw new NotSupportedException(entry.Diagnostic.Code + ": " + entry.Diagnostic.Message);
                ArchiveLimits.Add(ref selectedBytes, entry.Length, maxSelectedBytes, "maxSelectedBytes");
                entries[index] = entry;
                string cabinetKey = entry.CabinetName == null ? string.Empty : entry.CabinetName.Substring(1);
                var key = (cabinetKey, entry.FolderOrdinal!.Value);
                if (!groups.TryGetValue(key, out SelectedFolder? group)) {
                    group = new SelectedFolder(_cabinets[cabinetKey], entry.FolderExpandedBytes!.Value);
                    groups.Add(key, group); orderedGroups.Add(group);
                    ArchiveLimits.Add(ref expandedBytes, group.ExpandedBytes, _limits.Expanded - _readBytes, "MaxExpandedBytes");
                }
                group.Ordinals.Add(entry.CabinetOrdinal!.Value); group.ResultIndexes.Add(index);
            }
            cancellationToken.ThrowIfCancellationRequested();
            var payloads = new byte[count][];
            long workBefore = _limits.DecompressionWork.ConsumedBytes;
            try {
                foreach (SelectedFolder group in orderedGroups) {
                    cancellationToken.ThrowIfCancellationRequested();
                    long remaining = _limits.Expanded - _readBytes;
                    // Reserve only folders actually attempted. A later failure never exposes earlier payloads.
                    ArchiveLimits.Add(ref _readBytes, group.ExpandedBytes, _limits.Expanded, "MaxExpandedBytes");
                    byte[][] bytes = group.Cabinet.ReadFolderMembers(group.Ordinals, remaining, _limits.DecompressionWork, cancellationToken);
                    for (int index = 0; index < bytes.Length; index++) {
                        cancellationToken.ThrowIfCancellationRequested();
                        int resultIndex = group.ResultIndexes[index];
                        if (bytes[index].LongLength != entries[resultIndex].Length)
                            throw new InvalidDataException("Member length differs from its declared complete size.");
                        payloads[resultIndex] = bytes[index];
                    }
                }
            } catch (OfficeCabinetException exception) { throw _limits.Translate(exception); }
            var members = new List<ArchiveReadMember>(count);
            try {
                for (int index = 0; index < count; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    members.Add(new ArchiveReadMember(entries[index], new MemoryStream(payloads[index], writable: false)));
                }
                cancellationToken.ThrowIfCancellationRequested();
                return new ArchiveReadBatch(members, selectedBytes, expandedBytes, _limits.DecompressionWork.ConsumedBytes - workBefore);
            } catch {
                foreach (ArchiveReadMember member in members) member.Content.Dispose();
                throw;
            }
        }
    }

    private sealed class SelectedFolder {
        internal SelectedFolder(OfficeCabinetArchiveReader.Inspection cabinet, long expandedBytes) {
            Cabinet = cabinet; ExpandedBytes = expandedBytes;
        }
        internal OfficeCabinetArchiveReader.Inspection Cabinet { get; }
        internal long ExpandedBytes { get; }
        internal List<int> Ordinals { get; } = new();
        internal List<int> ResultIndexes { get; } = new();
    }
}
