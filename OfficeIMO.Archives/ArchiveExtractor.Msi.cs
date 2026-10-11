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
        var database = new MsiPackageReader(source, limits, token);
        IReadOnlyList<MsiPayloadFile> payload = database.ReadFiles();
        long decodedBudget = limits.Expanded;
        foreach (var group in payload.GroupBy(file => file.Cabinet, StringComparer.OrdinalIgnoreCase)) {
            token.ThrowIfCancellationRequested();
            if (decodedBudget == 0) throw new InvalidDataException("MSI cabinet expansion exceeds the configured aggregate byte limit.");
            var expected = group.ToDictionary(file => file.Id, StringComparer.OrdinalIgnoreCase);
            var found = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using Stream cabinet = database.OpenCabinet(group.Key);
            IReadOnlyList<OfficeCabinetEntry> entries = OfficeCabinetArchiveReader.Read(cabinet, limits.Input,
                decodedBudget, limits.Entry, limits.Entries, out long decodedBytes, token, name => {
                    if (!expected.TryGetValue(name, out MsiPayloadFile? file))
                        throw new InvalidDataException("MSI cabinet contains a member absent from its File/Media mapping.");
                    if (!found.Add(name)) throw new InvalidDataException("MSI cabinet repeats a File identifier.");
                    return file.Selected;
                });
            if (found.Count != expected.Count) throw new InvalidDataException("MSI cabinet is missing a declared File member.");
            foreach (OfficeCabinetEntry entry in entries) {
                MsiPayloadFile file = expected[entry.Name];
                if (entry.Data.LongLength != file.Length) throw new InvalidDataException("MSI cabinet member length differs from FileSize.");
                WriteBytes(staging, file.Path, entry.Data, limits, files, ref total, token);
            }
            decodedBudget -= decodedBytes;
        }
    }
}
