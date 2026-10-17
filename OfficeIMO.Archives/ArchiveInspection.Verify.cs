using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Archives {

    public sealed partial class ArchiveInspection {
        /// <summary>Verifies selected members, or all members when ordinals is null, without extracting files.
        /// CAB folders are decoded once per selection. Unsupported sources and corrupt groups receive structured outcomes.
        /// Cancellation and resource exhaustion are terminal. This operation does not establish authenticity.</summary>
        public ArchiveVerificationResult Verify(IReadOnlyList<int>? ordinals = null, CancellationToken cancellationToken = default) {
            lock (_gate) {
                ThrowIfDisposed(); cancellationToken.ThrowIfCancellationRequested();
                int count = ordinals?.Count ?? Entries.Count;
                if (count > Entries.Count) throw new ArgumentException("The request exceeds the inventory count.", nameof(ordinals));
                var selected = new ArchiveInspectionEntry[count];
                var seen = new HashSet<int>();
                var groups = new Dictionary<(string? Cabinet, int Folder), List<ArchiveInspectionEntry>>();
                for (int index = 0; index < count; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int ordinal = ordinals == null ? index : ordinals[index];
                    if (ordinal < 0 || ordinal >= Entries.Count) throw new ArgumentOutOfRangeException(nameof(ordinals));
                    if (!seen.Add(ordinal)) throw new ArgumentException("A source ordinal is repeated.", nameof(ordinals));
                    var entry = selected[index] = Entries[ordinal];
                    if (!entry.CanRead) continue;
                    var key = (entry.CabinetName, Format == ArchiveFormat.Zip ? ordinal : entry.FolderOrdinal!.Value);
                    if (!groups.TryGetValue(key, out var group)) groups.Add(key, group = new List<ArchiveInspectionEntry>());
                    group.Add(entry);
                }
                var outcomes = new Dictionary<int, ArchiveMemberVerification>();
                foreach (var entry in selected) if (!entry.CanRead)
                    outcomes.Add(entry.Ordinal, new ArchiveMemberVerification(entry, ArchiveVerificationStatus.Unsupported, entry.Diagnostic!.Code, entry.Diagnostic.Message));
                foreach (var group in groups.Values) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var indexes = new int[group.Count];
                    for (int index = 0; index < indexes.Length; index++) indexes[index] = group[index].Ordinal;
                    try {
                        if (Format == ArchiveFormat.Zip) VerifyZipMember(indexes[0], Stream.Null, cancellationToken);
                        else { using var batch = ReadMembers(indexes, long.MaxValue, cancellationToken); }
                        foreach (var entry in group) outcomes.Add(entry.Ordinal, new ArchiveMemberVerification(entry, ArchiveVerificationStatus.Verified, null, null));
                    } catch (InvalidDataException error) {
                        foreach (var entry in group) outcomes.Add(entry.Ordinal, new ArchiveMemberVerification(entry, ArchiveVerificationStatus.Invalid, "ARCHIVE_INTEGRITY", error.Message));
                    }
                }
                var result = new ArchiveMemberVerification[count];
                for (int index = 0; index < count; index++) result[index] = outcomes[selected[index].Ordinal];
                cancellationToken.ThrowIfCancellationRequested();
                return new ArchiveVerificationResult(Array.AsReadOnly(result), ResourceUsage);
            }
        }
    }
}
