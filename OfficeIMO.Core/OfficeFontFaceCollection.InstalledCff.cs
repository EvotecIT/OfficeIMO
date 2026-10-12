using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeFontFaceCollection {
    private readonly Dictionary<string, List<OfficeFontFace>> _installedCffSources = new(StringComparer.Ordinal);

    // The same bounded source reader and scoped decoder used by mathematical fonts
    // also serve ordinary CFF families. Paths and metadata remain discovery hints.
    internal bool TryAddInstalledCffFamily(string family, OfficeFontFaceDescriptor requested,
        IEnumerable<string> paths, int maximumDecodedBytes, long maximumSourceBytes,
        CancellationToken cancellationToken, out int decodedBytes, out string? error) {
        decodedBytes = 0;
        error = null;
        cancellationToken.ThrowIfCancellationRequested();
        string key = OfficeFontNameMetadata.FamilyKey(family);
        var visited = new HashSet<string>(StringComparer.Ordinal);
        int attempts = 0;
        bool nativeSelection = FontProgramProvider == null && FontVariationResolver == null;
        string? selectedPath = null;
        byte[]? selectedData = null;
        OfficeFontFaceDescriptor selectedDescriptor = default;
        bool selectedVariable = false;
        foreach (string path in paths) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!visited.Add(path)) continue;
            if (++attempts > 32) break;
            _installedCffSources.TryGetValue(path, out List<OfficeFontFace>? retained);
            using (var stream = nativeSelection && retained != null
                ? new System.IO.MemoryStream(retained[0].DataSnapshot, writable: false) : null) {
                if (!(stream == null ? OfficeFontNameMetadata.ReadFamilyKeys(path)
                    : OfficeFontNameMetadata.ReadFamilyKeys(stream)).Contains(key)) continue;
            }
            byte[]? data = nativeSelection && retained != null ? retained[0].DataSnapshot
                : ReadInstalledCffFont(path, maximumDecodedBytes, maximumSourceBytes, cancellationToken, out error);
            if (data == null) {
                if (error != null) return false;
                continue;
            }
            OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(data);
            if (reader == null || !reader.TryGetTable("CFF ", out _, out _) && !reader.TryGetTable("CFF2", out _, out _)) continue;
            bool variable = reader.TryGetTable("fvar", out _, out _);
            OfficeFontFaceDescriptor descriptor = reader.ReadFaceDescriptor();
            if (variable) descriptor = new OfficeFontFaceDescriptor(requested.Weight, descriptor.StretchPercent, descriptor.Slant);
            if (selectedData != null && OfficeFontFaceMatcher.Compare(descriptor, selectedDescriptor, requested) >= 0) continue;
            selectedPath = path;
            selectedData = data;
            selectedDescriptor = descriptor;
            selectedVariable = variable;
        }
        if (selectedData == null || selectedPath == null) return false;
        string resource = selectedDescriptor == OfficeFontFaceDescriptor.Regular ? family
            : CreateResourceFamilyName(family, selectedDescriptor, OfficeFontUnicodeRangeSet.All);
        if (nativeSelection && _installedCffSources.TryGetValue(selectedPath, out List<OfficeFontFace>? faces)) {
            OfficeFontFace? registered = faces.Find(face => face.Descriptor == selectedDescriptor);
            if (registered != null) {
                OfficeFontFace alias = registered.CreateAlias(family, resource);
                int index = _faces.FindLastIndex(face => face.FamilyName == family && face.ResourceFamilyName == resource);
                if (index >= 0) _faces[index] = alias;
                else _faces.Add(alias);
                return true;
            }
        }
        if (!TryAddCore(family, selectedData, selectedDescriptor.ToStyle(), selectedDescriptor,
                OfficeFontUnicodeRangeSet.All, resource, maximumDecodedBytes, out decodedBytes, out error,
                applyDescriptorWeight: selectedVariable)) return false;
        cancellationToken.ThrowIfCancellationRequested();
        if (nativeSelection) {
            if (!_installedCffSources.TryGetValue(selectedPath, out List<OfficeFontFace>? retained))
                _installedCffSources[selectedPath] = retained = new List<OfficeFontFace>();
            retained.Add(_faces.FindLast(face => face.FamilyName == family && face.ResourceFamilyName == resource)!);
        }
        return true;
    }
}
