using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Core.Internal.Cabinet;

namespace OfficeIMO.Archives;

public static partial class ArchiveExtractor {
    private static void ExtractMsi(Stream source, string staging, ArchiveLimits limits, List<ExtractedArchiveFile> files,
        ref long total, CancellationToken token) {
        using var database = new MsiPackageReader(source, limits, token);
        IReadOnlyList<MsiPayloadFile> payload = database.ReadFiles(out IReadOnlyList<string> cabinets);
        var groups = payload.ToLookup(file => file.Cabinet, StringComparer.Ordinal);
        if (cabinets.Count > limits.Cabinets) throw new InvalidDataException("MSI exceeds the configured embedded-cabinet count.");
        long cabinetInputBytes = 0;
        foreach (string name in cabinets) {
            ArchiveLimits.Add(ref cabinetInputBytes, database.GetCabinetLength(name), limits.Input, "Aggregate MSI cabinet input");
        }
        long decodedBudget = limits.Expanded;
        int remainingEntries = limits.Entries;
        using var aliases = new MsiAliasVerifier(staging, token);
        foreach (string name in cabinets) {
            token.ThrowIfCancellationRequested();
            var expected = groups[name].ToDictionary(file => file.Id, StringComparer.Ordinal);
            if (expected.Count > remainingEntries) throw new InvalidDataException("MSI cabinets exceed the configured aggregate member count.");
            remainingEntries -= expected.Count;
            var found = new HashSet<string>(StringComparer.Ordinal);
            int previousSequence = 0;
            using Stream cabinet = database.OpenCabinet(name);
            // A Media-only cabinet may be empty; any actual member is rejected by the validator.
            IReadOnlyList<OfficeCabinetEntry> entries = OfficeCabinetArchiveReader.Read(cabinet, limits.Input,
                decodedBudget, limits.Entry, Math.Max(1, expected.Count), out long decodedBytes, token, memberName => {
                    if (!expected.TryGetValue(memberName, out MsiPayloadFile? file))
                        throw new InvalidDataException("MSI cabinet contains a member absent from its File/Media mapping.");
                    if (!found.Add(memberName)) throw new InvalidDataException("MSI cabinet repeats a File identifier.");
                    return file.Target.Selected || file.Target.AliasCount > 1;
                }, limits.DecompressionWork, (memberName, length) => {
                    if (!expected.TryGetValue(memberName, out MsiPayloadFile? file))
                        throw new InvalidDataException("MSI cabinet contains a member absent from its File/Media mapping.");
                    if (length != file.Length) throw new InvalidDataException("MSI cabinet member length differs from FileSize.");
                    // The reader visits every CFFILE in authored order before selection or decoding.
                    if (file.Sequence <= previousSequence)
                        throw new InvalidDataException("MSI cabinet member order differs from File sequence order.");
                    previousSequence = file.Sequence;
                });
            if (found.Count != expected.Count) throw new InvalidDataException("MSI cabinet is missing a declared File member.");
            foreach (OfficeCabinetEntry entry in entries) {
                MsiPayloadFile file = expected[entry.Name];
                if (entry.Data.LongLength != file.Length) throw new InvalidDataException("MSI cabinet member length differs from FileSize.");
                bool first = file.Target.AliasCount == 1 || aliases.Verify(file.Target, entry.Data);
                if (first && file.Target.Selected)
                    WriteBytes(staging, file.Target.Path, entry.Data, limits, files, ref total, token);
            }
            decodedBudget -= decodedBytes;
        }
    }
}
