using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.Cabinet;

namespace OfficeIMO.Archives;

public sealed partial class ArchiveInspection {
    private IReadOnlyList<ArchiveInspectionEntry> ReadMsiInventory(byte[] input, OfficeCabinetDirectoryBudget budget,
        int directoryRecords, CancellationToken token) {
        _databaseSource = new MemoryStream(input, writable: false);
        _database = new MsiPackageReader(_databaseSource, _limits, token, Math.Min(65536, directoryRecords));
        IReadOnlyList<OfficeCompoundFileEntry> streams = _database.StorageStreams;
        budget.Consume(streams.Count);
        StorageStreams = Array.AsReadOnly(streams.Select((stream, index) =>
            new ArchiveStorageStream(index, stream.DirectoryOrder, stream.Name, stream.Size)).ToArray());
        IReadOnlyList<MsiPayloadFile> files = _database.ReadFiles(out IReadOnlyList<string> cabinetNames, inspection: true);
        if (cabinetNames.Count > _limits.Cabinets) throw new InvalidDataException("MSI exceeds the configured embedded-cabinet count.");
        long declaredFileBytes = 0;
        foreach (MsiPayloadFile file in files) ArchiveLimits.Add(ref declaredFileBytes, file.Length, _limits.Expanded, "Declared MSI member content");
        long cabinetInput = 0;
        long declaredFolderBytes = 0;
        var mappings = new Dictionary<string, OfficeCabinetMemberInfo>(StringComparer.Ordinal);
        var groups = files.Where(file => file.RawCabinet.StartsWith("#", StringComparison.Ordinal))
            .ToLookup(file => file.Cabinet, StringComparer.Ordinal);
        int actualEntries = 0;
        foreach (string name in cabinetNames) {
            token.ThrowIfCancellationRequested();
            long length = _database.GetCabinetLength(name);
            ArchiveLimits.Add(ref cabinetInput, length, _limits.Input, "Aggregate MSI cabinet input");
            using Stream source = _database.OpenCabinet(name);
            byte[] bytes = OfficeArchiveSafety.ReadEntryBytes(source, length, _limits.Input, token);
            var cabinet = OfficeCabinetArchiveReader.Inspect(bytes, _limits.Expanded - declaredFolderBytes,
                _limits.Entry, Math.Max(1, _limits.Entries - actualEntries), budget, token);
            _cabinets.Add(name, cabinet);
            actualEntries = checked(actualEntries + cabinet.Members.Count);
            if (actualEntries > _limits.Entries) throw new InvalidDataException("MSI cabinets exceed the configured aggregate member count.");
            ArchiveLimits.Add(ref declaredFolderBytes, cabinet.ExpandedBytes, _limits.Expanded, "Declared MSI folder expansion");
            var expected = groups[name].ToDictionary(file => file.Id, StringComparer.Ordinal);
            var found = new HashSet<string>(StringComparer.Ordinal);
            int previousSequence = 0;
            foreach (OfficeCabinetMemberInfo member in cabinet.Members) {
                token.ThrowIfCancellationRequested();
                if (!expected.TryGetValue(member.Name, out MsiPayloadFile? file))
                    throw new InvalidDataException("MSI cabinet contains a member absent from its File/Media mapping.");
                if (!found.Add(member.Name)) throw new InvalidDataException("MSI cabinet repeats a File identifier.");
                if (member.Length != file.Length) throw new InvalidDataException("MSI cabinet member length differs from FileSize.");
                if (file.Sequence <= previousSequence) throw new InvalidDataException("MSI cabinet member order differs from File sequence order.");
                previousSequence = file.Sequence;
                if (!file.LooseSource) mappings.Add(file.Id, member);
            }
            if (expected.Values.Any(file => !file.LooseSource && !found.Contains(file.Id)))
                throw new InvalidDataException("MSI cabinet is missing a declared File member.");
        }
        var result = new List<ArchiveInspectionEntry>(files.Count);
        for (int index = 0; index < files.Count; index++) {
            token.ThrowIfCancellationRequested();
            MsiPayloadFile file = files[index];
            if (mappings.TryGetValue(file.Id, out OfficeCabinetMemberInfo? member))
                result.Add(Entry(member, index, file.RawFileName, file.Path, file.RawCabinet, file.Sequence));
            else {
                var diagnostic = new ArchiveInspectionDiagnostic(file.LooseSource ? "MSI_LOOSE_SOURCE" : "MSI_EXTERNAL_CABINET",
                    file.LooseSource ? "Loose MSI source files are inventoried without accessing their source paths."
                        : "External MSI cabinets are inventoried without accessing their paths or URLs.");
                result.Add(new ArchiveInspectionEntry(index, file.Id, file.RawFileName, file.Path, file.Length,
                    file.RawCabinet, null, file.Sequence, ArchiveCompression.None, null, null, null, null, null, null, diagnostic));
            }
        }
        return result.AsReadOnly();
    }
}
