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
        IReadOnlyList<MsiPayloadFile> payload = database.ReadFiles();
        var groups = payload.GroupBy(file => file.Cabinet, StringComparer.Ordinal).ToArray();
        if (groups.Length > limits.Cabinets) throw new InvalidDataException("MSI exceeds the configured embedded-cabinet count.");
        long cabinetInputBytes = 0;
        foreach (var group in groups) {
            ArchiveLimits.Add(ref cabinetInputBytes, database.GetCabinetLength(group.Key), limits.Input, "Aggregate MSI cabinet input");
        }
        long decodedBudget = limits.Expanded;
        int remainingEntries = limits.Entries;
        using var aliases = new MsiAliasVerifier(staging, token);
        foreach (var group in groups) {
            token.ThrowIfCancellationRequested();
            var expected = group.ToDictionary(file => file.Id, StringComparer.OrdinalIgnoreCase);
            if (expected.Count > remainingEntries) throw new InvalidDataException("MSI cabinets exceed the configured aggregate member count.");
            remainingEntries -= expected.Count;
            var found = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using Stream cabinet = database.OpenCabinet(group.Key);
            IReadOnlyList<OfficeCabinetEntry> entries = OfficeCabinetArchiveReader.Read(cabinet, limits.Input,
                decodedBudget, limits.Entry, expected.Count, out long decodedBytes, token, name => {
                    if (!expected.TryGetValue(name, out MsiPayloadFile? file))
                        throw new InvalidDataException("MSI cabinet contains a member absent from its File/Media mapping.");
                    if (!found.Add(name)) throw new InvalidDataException("MSI cabinet repeats a File identifier.");
                    return file.Target.Selected || file.Target.AliasCount > 1;
                }, limits.DecompressionWork, (name, length) => {
                    if (!expected.TryGetValue(name, out MsiPayloadFile? file))
                        throw new InvalidDataException("MSI cabinet contains a member absent from its File/Media mapping.");
                    if (length != file.Length) throw new InvalidDataException("MSI cabinet member length differs from FileSize.");
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
