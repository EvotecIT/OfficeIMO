using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OfficeIMO.Archives;

internal sealed partial class MsiPackageReader {
    internal IReadOnlyList<MsiPayloadFile> ReadFiles(out IReadOnlyList<string> cabinets) {
        var directories = Index(Table("Directory", "Directory", "Directory_Parent", "DefaultDir"), "Directory");
        var components = Index(Table("Component", "Component", "ComponentId", "Directory_", "Attributes", "Condition", "KeyPath"), "Component");
        object?[][] fileRows = Table("File", "File", "Component_", "FileName", "FileSize", "Version", "Language", "Attributes", "Sequence");
        _ = Index(fileRows, "File", StringComparer.OrdinalIgnoreCase);
        var mediaRows = Table("Media", "DiskId", "LastSequence", "DiskPrompt", "Cabinet", "VolumeLabel", "Source");
        foreach (object?[] row in directories.Values)
            if (row[1] != null) RequiredIdentifier(row[1], "Directory.Directory_Parent");
        foreach (object?[] row in components.Values) {
            RequiredIdentifier(row[2], "Component.Directory_");
            if (row[5] != null) RequiredIdentifier(row[5], "Component.KeyPath");
        }
        foreach (object?[] row in fileRows) RequiredIdentifier(row[1], "File.Component_");
        if (fileRows.Length > _limits.Entries) throw new InvalidDataException("MSI exceeds the configured file count.");
        bool compressedByDefault = ReadDefaultCompression();
        foreach (object?[] row in fileRows) {
            _token.ThrowIfCancellationRequested();
            int attributes = row[6] is int flags ? flags : 0;
            int compression = attributes & 0x6000;
            if (compression == 0x6000) throw new InvalidDataException("MSI File compression overrides conflict.");
            if (compression == 0x2000 || compression == 0 && !compressedByDefault)
                throw new NotSupportedException("MSI loose uncompressed source files are unsupported.");
        }
        var media = new List<Tuple<int, int, string>>();
        var disks = new HashSet<int>();
        foreach (object?[] row in mediaRows) {
            int disk = RequiredInteger(row[0], "Media.DiskId");
            int last = RequiredInteger(row[1], "Media.LastSequence");
            string cabinet = RequiredString(row[3], "Media.Cabinet");
            if (disk <= 0 || last < 0 || !disks.Add(disk)) throw new InvalidDataException("MSI media identifiers or sequences are invalid.");
            if (!cabinet.StartsWith("#", StringComparison.Ordinal) || cabinet.Length == 1)
                throw new NotSupportedException("MSI external cabinets and loose source files are unsupported.");
            if (!_streams.ContainsKey(EncodeStreamName(cabinet.Substring(1), table: false))) throw new InvalidDataException("An embedded MSI cabinet is absent.");
            media.Add(Tuple.Create(disk, last, cabinet.Substring(1)));
        }
        media = media.OrderBy(item => item.Item1).ToList();
        if (media.Count != 0 && media[0].Item1 != 1)
            throw new InvalidDataException("MSI first media DiskId must be 1.");
        int previousSequence = -1;
        foreach (var disk in media) {
            if (disk.Item2 < previousSequence) throw new InvalidDataException("MSI media sequence boundaries move backwards in disk order.");
            previousSequence = disk.Item2;
        }
        cabinets = media.Select(item => item.Item3).Distinct(StringComparer.Ordinal).ToArray();
        var sequences = new HashSet<int>();
        var directoryCache = new Dictionary<string, MsiResolvedDirectory>(StringComparer.Ordinal);
        var roots = directories.Where(pair => pair.Value[1] == null || string.Equals(pair.Key, pair.Value[1] as string, StringComparison.Ordinal)).Take(2).ToArray();
        if (roots.Length != 1 || !string.Equals(roots[0].Key, "TARGETDIR", StringComparison.Ordinal) ||
            !string.Equals(roots[0].Value[2] as string, "SourceDir", StringComparison.Ordinal))
            throw new InvalidDataException("MSI requires exactly one TARGETDIR root with SourceDir as DefaultDir.");
        foreach (string directoryId in directories.Keys)
            ResolveDirectory(directoryId, directories, directoryCache);
        var result = new List<MsiPayloadFile>();
        foreach (object?[] row in fileRows) {
            _token.ThrowIfCancellationRequested();
            string id = RequiredIdentifier(row[0], "File.File");
            string component = RequiredIdentifier(row[1], "File.Component_");
            if (!components.TryGetValue(component, out object?[]? owner)) throw new InvalidDataException("MSI file references an absent component.");
            MsiResolvedDirectory resolved = ResolveDirectory(RequiredIdentifier(owner[2], "Component.Directory_"), directories, directoryCache);
            string directory = resolved.Path;
            string filename = InstalledName(LongName(RequiredString(row[2], "File.FileName")));
            ArchivePaths.ValidateSegment(filename);
            long pathLength = (long)directory.Length + (directory.Length == 0 ? 0 : 1) + filename.Length;
            if (resolved.PathSegments >= _limits.Depth || pathLength > _limits.PathLength)
                throw new InvalidDataException("MSI authored file path exceeds configured bounds.");
            // A pool string may be reused by many File rows. Check bounds before
            // allocating each target; only collision registration waits for aliases.
            string path = directory.Length == 0 ? filename : directory + "/" + filename;
            int length = RequiredInteger(row[3], "File.FileSize");
            _limits.CheckEntry(length);
            int sequence = RequiredInteger(row[7], "File.Sequence");
            if (sequence <= 0 || !sequences.Add(sequence)) throw new InvalidDataException("MSI file sequences are invalid or duplicated.");
            string cabinet = FindCabinet(media, sequence);
            result.Add(new MsiPayloadFile(id, path, length, cabinet, sequence));
        }
        PlanTargets(result);
        return result;
    }
    private static string FindCabinet(List<Tuple<int, int, string>> media, int sequence) {
        int low = 0, high = media.Count - 1, match = -1;
        while (low <= high) {
            int middle = low + (high - low) / 2;
            if (sequence <= media[middle].Item2) { match = middle; high = middle - 1; }
            else low = middle + 1;
        }
        if (match < 0) throw new InvalidDataException("MSI file has no corresponding media sequence range.");
        return media[match].Item3;
    }
    private static string LongName(string value) {
        // Both halves are bounded before splitting, even when the schema permits unbounded strings.
        if (value.Length > 2 * 255 + 1)
            throw new InvalidDataException("MSI authored short/long filename exceeds supported bounds.");
        string[] names = value.Split('|');
        if (names.Length > 2 || names[names.Length - 1].Length == 0) throw new InvalidDataException("MSI contains an invalid short/long filename.");
        return names[names.Length - 1];
    }
}

internal sealed class MsiPayloadFile {
    internal MsiPayloadFile(string id, string path, long length, string cabinet, int sequence) { Id = id; Path = path; Length = length; Cabinet = cabinet; Sequence = sequence; }
    internal string Id { get; }
    internal string Path { get; }
    internal long Length { get; }
    internal string Cabinet { get; }
    internal int Sequence { get; }
    internal MsiPayloadTarget Target { get; set; } = null!;
}
