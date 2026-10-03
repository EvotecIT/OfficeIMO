using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeFontFaceCollection {
    private OfficeFontFace? ResolveMathematicalFace(string? text, string family, OfficeFontStyle style) {
        foreach ((OfficeFontFace face, _) in ResolveFamilyCandidates(family, OfficeFontFaceDescriptor.FromStyle(style))) {
            bool explicitResource = string.Equals(face.ResourceFamilyName, family, StringComparison.OrdinalIgnoreCase);
            if (text == null || CoversPlannedText(face, text, requireUnicodeRange: !explicitResource)) return face;
        }
        return null;
    }

    private readonly Dictionary<string, List<OfficeFontFace>> _installedMathCffSources = new(StringComparer.Ordinal);

    private bool TryAddInstalledMathematicalFamily(OfficeFontFaceDescriptor requested, int maximumDecodedBytes,
        long maximumSourceBytes, CancellationToken cancellationToken, out int decodedBytes, out string? error) {
        decodedBytes = 0;
        error = null;
        foreach (string family in OfficeSystemFontFamilyAliases.Expand("math")) {
            if (OfficeSystemFontFamilyAliases.IsMath(family)) continue;
            // A retained TrueType source is already charged to this operation. Do not
            // spend its remaining budget probing the same file as a possible CFF source.
            bool retainedTrueType = FontProgramProvider == null && FontVariationResolver == null
                && System.Linq.Enumerable.Any(_installedSourceFaces.Values, face =>
                    !face.Program.IsOpenTypeCff
                    && string.Equals(face.FamilyName, family, StringComparison.OrdinalIgnoreCase));
            if (retainedTrueType && TryAddInstalledFamily(family, requested, maximumDecodedBytes,
                    cancellationToken, out decodedBytes, out error, maximumSourceBytes)) {
                AddAlias("math", family);
                return true;
            }
            if (error != null) return false;
            // Probe both supported formats before advancing to the next family. Start
            // with the bounded source read; TrueType collections use named-face extraction.
            if (TryAddInstalledMathCff(family, requested, maximumDecodedBytes, maximumSourceBytes,
                    cancellationToken, out decodedBytes, out error)) return true;
            if (error != null) return false;
            if (TryAddInstalledFamily(family, requested, maximumDecodedBytes, cancellationToken,
                    out decodedBytes, out error, maximumSourceBytes)) {
                AddAlias("math", family);
                return true;
            }
            if (error != null) return false;
        }
        return false;
    }

    // Mathematical installed fonts commonly use CFF. Register them through the existing
    // bounded scoped loader so measurement, outlines and PDF output share the same face.
    private bool TryAddInstalledMathCff(string family, OfficeFontFaceDescriptor requested, int maximumDecodedBytes,
        long maximumSourceBytes, CancellationToken cancellationToken, out int decodedBytes, out string? error) {
        decodedBytes = 0;
        error = null;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        int attempts = 0;
        bool nativeSelection = FontProgramProvider == null && FontVariationResolver == null;
        foreach (string path in OfficeTrueTypeFont.CandidateFamilyPaths(family)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!visited.Add(path)) continue;
            if (++attempts > 32) return false;
            List<OfficeFontFace>? retained = null;
            if (nativeSelection) _installedMathCffSources.TryGetValue(path, out retained);
            byte[]? data = retained != null ? retained[0].DataSnapshot
                : ReadInstalledMathFont(path, maximumDecodedBytes, maximumSourceBytes, cancellationToken, out error);
            if (data == null) {
                if (error != null) return false;
                continue;
            }
            OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(data);
            if (reader == null || !reader.TryGetTable("MATH", out _, out _)
                || !reader.TryGetTable("CFF ", out _, out _) && !reader.TryGetTable("CFF2", out _, out _)
                || !string.Equals(reader.ReadFamilyName(), family, StringComparison.OrdinalIgnoreCase)) continue;
            bool variable = reader.TryGetTable("fvar", out _, out _);
            OfficeFontFaceDescriptor descriptor = variable
                ? new OfficeFontFaceDescriptor(requested.Weight, 100D, OfficeFontSlant.Normal)
                : OfficeFontFaceDescriptor.Regular;
            OfficeFontFace? registered = retained?.Find(face => face.Descriptor == descriptor);
            if (registered != null) {
                int index = _faces.FindLastIndex(face => face.FamilyName == "math"
                    && face.ResourceFamilyName == registered.ResourceFamilyName);
                if (index >= 0) _faces[index] = registered;
                else _faces.Add(registered);
                return true;
            }
            // Static mathematical programs stay regular; synthetic styling remains the
            // existing paint contract. Variable weights get distinct export resources.
            string resource = descriptor == OfficeFontFaceDescriptor.Regular ? "math"
                : CreateResourceFamilyName("math", descriptor, OfficeFontUnicodeRangeSet.All);
            // Another variable weight creates its own CFF reader and shaping snapshots.
            // Reusing source bytes must not exempt those new buffers from accounting.
            if (!TryAddCore("math", data, descriptor.ToStyle(), descriptor, OfficeFontUnicodeRangeSet.All,
                    resource, maximumDecodedBytes, out decodedBytes, out error, applyDescriptorWeight: variable)) return false;
            cancellationToken.ThrowIfCancellationRequested();
            if (nativeSelection) {
                if (retained == null) _installedMathCffSources[path] = retained = new List<OfficeFontFace>();
                retained.Add(_faces.FindLast(face => face.FamilyName == "math" && face.ResourceFamilyName == resource)!);
            }
            return true;
        }
        return false;
    }

    private static byte[]? ReadInstalledMathFont(string path, int maximumDecodedBytes,
        long maximumSourceBytes, CancellationToken cancellationToken, out string? error) {
        error = null;
        try {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            long length = stream.Length;
            // TrueType and collection candidates belong to the named-face loader below.
            // Do not spend a CFF snapshot budget on a file this route cannot accept.
            if (length < 4 || stream.ReadByte() != 'O' || stream.ReadByte() != 'T'
                || stream.ReadByte() != 'T' || stream.ReadByte() != 'O') return null;
            cancellationToken.ThrowIfCancellationRequested();
            stream.Position = 0;
            if (length > maximumSourceBytes) {
                error = "Installed font source exceeds the per-resource byte limit.";
                return null;
            }
            // Keep source and loader-owned decoded snapshots inside the same operation budget.
            if (length > maximumDecodedBytes / 2 || length > int.MaxValue) {
                error = "Installed font source exceeds the byte limit.";
                return null;
            }
            if (length == 0) return null;
            var data = new byte[(int)length];
            int offset = 0;
            while (offset < data.Length) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = stream.Read(data, offset, Math.Min(65536, data.Length - offset));
                if (read == 0) return null;
                offset += read;
            }
            cancellationToken.ThrowIfCancellationRequested();
            return data;
        } catch (IOException) {
            return null;
        } catch (UnauthorizedAccessException) {
            return null;
        }
    }
}
