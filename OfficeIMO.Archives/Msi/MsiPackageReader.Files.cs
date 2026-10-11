using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OfficeIMO.Archives;

internal sealed partial class MsiPackageReader {
    internal IReadOnlyList<MsiPayloadFile> ReadFiles() {
        var directories = Index(Table("Directory", "Directory", "Directory_Parent", "DefaultDir"), "Directory");
        var components = Index(Table("Component", "Component", "ComponentId", "Directory_", "Attributes", "Condition", "KeyPath"), "Component");
        object?[][] fileRows = Table("File", "File", "Component_", "FileName", "FileSize", "Version", "Language", "Attributes", "Sequence");
        _ = Index(fileRows, "File", StringComparer.OrdinalIgnoreCase);
        var mediaRows = Table("Media", "DiskId", "LastSequence", "DiskPrompt", "Cabinet", "VolumeLabel", "Source");
        if (fileRows.Length > _limits.Entries) throw new InvalidDataException("MSI exceeds the configured file count.");
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
        int previousSequence = -1;
        foreach (var disk in media) {
            if (disk.Item2 < previousSequence) throw new InvalidDataException("MSI media sequence boundaries move backwards in disk order.");
            previousSequence = disk.Item2;
        }
        var paths = new ArchivePaths(_limits);
        var sequences = new HashSet<int>();
        var directoryCache = new Dictionary<string, MsiResolvedDirectory>(StringComparer.Ordinal);
        if (directories.Count(pair => pair.Value[1] == null || string.Equals(pair.Key, pair.Value[1] as string, StringComparison.Ordinal)) != 1)
            throw new InvalidDataException("MSI requires exactly one authored directory root.");
        foreach (string directoryId in directories.Keys)
            ResolveDirectory(directoryId, directories, directoryCache);
        var result = new List<MsiPayloadFile>();
        long declared = 0;
        foreach (object?[] row in fileRows) {
            _token.ThrowIfCancellationRequested();
            string id = RequiredString(row[0], "File.File");
            string component = RequiredString(row[1], "File.Component_");
            if (!components.TryGetValue(component, out object?[]? owner)) throw new InvalidDataException("MSI file references an absent component.");
            string directory = ResolveDirectory(RequiredString(owner[2], "Component.Directory_"), directories, directoryCache).Path;
            string filename = LongName(RequiredString(row[2], "File.FileName"));
            if (filename.IndexOf('/') >= 0 || filename.IndexOf('\\') >= 0) throw new InvalidDataException("MSI FileName contains directory separators.");
            string path = paths.Add(directory.Length == 0 ? filename : directory + "/" + filename);
            int length = RequiredInteger(row[3], "File.FileSize");
            _limits.CheckEntry(length);
            int attributes = row[6] is int flags ? flags : 0;
            if ((attributes & 0x2000) != 0) throw new NotSupportedException("MSI loose uncompressed source files are unsupported.");
            int sequence = RequiredInteger(row[7], "File.Sequence");
            if (sequence <= 0 || !sequences.Add(sequence)) throw new InvalidDataException("MSI file sequences are invalid or duplicated.");
            string cabinet = FindCabinet(media, sequence);
            bool selected = _limits.Select(path);
            if (selected) ArchiveLimits.Add(ref declared, length, _limits.Expanded, "Selected MSI content");
            result.Add(new MsiPayloadFile(id, path, length, cabinet, selected));
        }
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
        string[] names = value.Split('|');
        if (names.Length > 2 || names[names.Length - 1].Length == 0) throw new InvalidDataException("MSI contains an invalid short/long filename.");
        return names[names.Length - 1];
    }
}

internal sealed class MsiPayloadFile {
    internal MsiPayloadFile(string id, string path, long length, string cabinet, bool selected) { Id = id; Path = path; Length = length; Cabinet = cabinet; Selected = selected; }
    internal string Id { get; }
    internal string Path { get; }
    internal long Length { get; }
    internal string Cabinet { get; }
    internal bool Selected { get; }
}
