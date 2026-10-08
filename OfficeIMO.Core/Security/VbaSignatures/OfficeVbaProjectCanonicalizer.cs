using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Security;

/// <summary>Bounded managed implementation of the MS-OVBA signature-binding transcripts.</summary>
internal static partial class OfficeVbaProjectCanonicalizer {
    private static readonly byte[][] V3DefaultAttributes = {
        Encoding.ASCII.GetBytes("Attribute VB_Base = \"0{00020820-0000-0000-C000-000000000046}\""),
        Encoding.ASCII.GetBytes("Attribute VB_GlobalNameSpace = False"),
        Encoding.ASCII.GetBytes("Attribute VB_Creatable = False"),
        Encoding.ASCII.GetBytes("Attribute VB_PredeclaredId = True"),
        Encoding.ASCII.GetBytes("Attribute VB_Exposed = True"),
        Encoding.ASCII.GetBytes("Attribute VB_TemplateDerived = False"),
        Encoding.ASCII.GetBytes("Attribute VB_Customizable = True")
    };

    internal sealed class Result {
        internal Result(byte[] contentNormalizedData, byte[] formsNormalizedData,
            byte[] v3ContentNormalizedData, byte[] projectNormalizedData) {
            ContentNormalizedData = contentNormalizedData;
            FormsNormalizedData = formsNormalizedData;
            V3ContentNormalizedData = v3ContentNormalizedData;
            ProjectNormalizedData = projectNormalizedData;
        }

        internal byte[] ContentNormalizedData { get; }
        internal byte[] FormsNormalizedData { get; }
        internal byte[] V3ContentNormalizedData { get; }
        internal byte[] ProjectNormalizedData { get; }

        internal byte[] ComputeLegacyHash() => Hash(ContentNormalizedData, false);

        internal byte[] ComputeLegacySipHash() => Hash(ContentNormalizedData, true);

        internal byte[] ComputeAgileHash() => Hash(Concat(ContentNormalizedData, FormsNormalizedData), true);

        internal byte[] ComputeV3Hash() => Hash(Concat(V3ContentNormalizedData, ProjectNormalizedData), true);

        private static byte[] Hash(byte[] bytes, bool sha256) {
            using HashAlgorithm algorithm = sha256 ? SHA256.Create() : MD5.Create();
            return algorithm.ComputeHash(bytes);
        }
    }

    internal static bool TryCreate(byte[] projectBytes, long maximumExpandedBytes,
        out Result? result, out string detail) {
        result = null;
        detail = string.Empty;
        if (projectBytes == null || projectBytes.Length == 0) {
            detail = "The VBA project is empty.";
            return false;
        }
        if (maximumExpandedBytes <= 0 || maximumExpandedBytes > int.MaxValue) {
            detail = "The VBA canonicalization byte limit is invalid.";
            return false;
        }
        if (!OfficeCompoundFileReader.TryRead(projectBytes, out OfficeCompoundFile? compound, out string? compoundError)
            || compound == null) {
            detail = compoundError ?? "The VBA project is not a valid compound file.";
            return false;
        }
        if (!compound.Streams.TryGetValue("VBA/dir", out byte[]? compressedDirectory)) {
            detail = "The VBA project has no VBA/dir stream.";
            return false;
        }
        if (!TryDecompress(compressedDirectory, checked((int)maximumExpandedBytes), out byte[] directory, out detail)) {
            return false;
        }
        if (!DirectoryModel.TryParse(directory, checked((int)maximumExpandedBytes),
                out DirectoryModel? model, out detail) || model == null) {
            return false;
        }
        if (!TryBuildContentNormalizedData(compound, model, checked((int)maximumExpandedBytes),
            out byte[] content, out detail)) {
            return false;
        }
        if (!TryBuildFormsNormalizedData(compound, model, checked((int)maximumExpandedBytes),
            out byte[] forms, out detail)) {
            return false;
        }
        if (!TryBuildV3ContentNormalizedData(compound, model, checked((int)maximumExpandedBytes),
            out byte[] v3, out detail)) {
            return false;
        }
        if (!TryBuildProjectNormalizedData(compound, model, checked((int)maximumExpandedBytes),
            out byte[] project, out detail)) {
            return false;
        }
        if ((long)content.Length + forms.Length + v3.Length + project.Length > maximumExpandedBytes) {
            detail = "The aggregate VBA canonicalization transcript exceeds the configured byte limit.";
            return false;
        }
        result = new Result(content, forms, v3, project);
        return true;
    }

    private static bool TryBuildContentNormalizedData(OfficeCompoundFile compound, DirectoryModel model,
        int maximumBytes, out byte[] output, out string detail) {
        var buffer = new BoundedBuffer(maximumBytes);
        if (!buffer.TryAppend(model.ProjectName) || !buffer.TryAppend(model.ProjectConstants)) {
            output = Array.Empty<byte>();
            detail = "The VBA content-normalized transcript exceeds the configured byte limit.";
            return false;
        }
        foreach (ReferenceModel reference in model.References) {
            byte[] normalized = reference.Normalize();
            if (!buffer.TryAppend(normalized)) {
                output = Array.Empty<byte>();
                detail = "The VBA reference transcript exceeds the configured byte limit.";
                return false;
            }
        }
        foreach (ModuleModel module in model.Modules) {
            if (!TryReadModuleSource(compound, module, maximumBytes, out byte[] source, out detail)) {
                output = Array.Empty<byte>();
                return false;
            }
            foreach (byte[] line in SplitLines(source)) {
                if (!StartsWithAsciiIgnoreCase(line, "attribute") && !buffer.TryAppend(line)) {
                    output = Array.Empty<byte>();
                    detail = "The VBA module transcript exceeds the configured byte limit.";
                    return false;
                }
            }
        }
        output = buffer.ToArray();
        detail = string.Empty;
        return true;
    }

    private static bool TryBuildV3ContentNormalizedData(OfficeCompoundFile compound, DirectoryModel model,
        int maximumBytes, out byte[] output, out string detail) {
        var buffer = new BoundedBuffer(maximumBytes);
        if (!buffer.TryAppend(model.V3Prefix) || !buffer.TryAppend(model.ProjectModulesHeader)
            || !buffer.TryAppend(model.ProjectCookieHeader)) {
            output = Array.Empty<byte>();
            detail = "The VBA V3 project transcript exceeds the configured byte limit.";
            return false;
        }
        foreach (ModuleModel module in model.Modules) {
            if ((module.TypeId == 0x0021 && !buffer.TryAppend(module.TypeRecord))
                || (module.ReadOnlyRecord != null && !buffer.TryAppend(module.ReadOnlyRecord))
                || (module.PrivateRecord != null && !buffer.TryAppend(module.PrivateRecord))) {
                output = Array.Empty<byte>();
                detail = "The VBA V3 module metadata exceeds the configured byte limit.";
                return false;
            }
            if (!TryReadModuleSource(compound, module, maximumBytes, out byte[] source, out detail)) {
                output = Array.Empty<byte>();
                return false;
            }
            bool hashModuleName = false;
            foreach (byte[] line in SplitLinesWithoutTerminalEmptyLine(source)) {
                bool attribute = StartsWithAsciiIgnoreCase(line, "attribute");
                if (attribute && StartsWithAsciiIgnoreCase(line, "Attribute VB_Name = ")) continue;
                if (attribute && V3DefaultAttributes.Any(defaultValue => BytesEqual(line, defaultValue))) continue;
                if (!buffer.TryAppend(line) || !buffer.TryAppendByte(0x0A)) {
                    output = Array.Empty<byte>();
                    detail = "The VBA V3 source transcript exceeds the configured byte limit.";
                    return false;
                }
                hashModuleName = true;
            }
            if (hashModuleName && (!buffer.TryAppend(module.PreferredNameBytes) || !buffer.TryAppendByte(0x0A))) {
                output = Array.Empty<byte>();
                detail = "The VBA V3 module-name transcript exceeds the configured byte limit.";
                return false;
            }
        }
        if (!buffer.TryAppend(model.TerminatorRecord)) {
            output = Array.Empty<byte>();
            detail = "The VBA V3 transcript exceeds the configured byte limit.";
            return false;
        }
        output = buffer.ToArray();
        detail = string.Empty;
        return true;
    }

    private static IEnumerable<byte[]> SplitLinesWithoutTerminalEmptyLine(byte[] bytes) {
        byte[]? pending = null;
        foreach (byte[] line in SplitLines(bytes)) {
            if (pending != null) yield return pending;
            pending = line;
        }
        if (pending is { Length: > 0 }) yield return pending;
    }

    private static bool TryBuildFormsNormalizedData(OfficeCompoundFile compound, DirectoryModel model,
        int maximumBytes, out byte[] output, out string detail) {
        output = Array.Empty<byte>();
        if (!compound.Streams.TryGetValue("PROJECT", out byte[]? project)) {
            detail = string.Empty;
            return true;
        }
        var buffer = new BoundedBuffer(maximumBytes);
        foreach (ProjectProperty property in ReadProjectProperties(project)) {
            if (!AsciiEquals(property.Name, "BaseClass")) continue;
            ModuleModel? module = model.Modules.FirstOrDefault(item =>
                BytesEqualAsciiIgnoreCase(item.AnsiName, property.Value));
            if (module == null) continue;
            if (!TryAppendDesignerStorage(compound, module.StreamName, buffer)) {
                detail = "The VBA forms-normalized transcript exceeds the configured byte limit.";
                return false;
            }
        }
        output = buffer.ToArray();
        detail = string.Empty;
        return true;
    }

    private static bool TryBuildProjectNormalizedData(OfficeCompoundFile compound, DirectoryModel model,
        int maximumBytes, out byte[] output, out string detail) {
        output = Array.Empty<byte>();
        if (!compound.Streams.TryGetValue("PROJECT", out byte[]? project)) {
            detail = string.Empty;
            return true;
        }
        var buffer = new BoundedBuffer(maximumBytes);
        foreach (ProjectProperty property in ReadProjectProperties(project)) {
            if (AsciiEquals(property.Name, "BaseClass")) {
                ModuleModel? module = model.Modules.FirstOrDefault(item =>
                    BytesEqualAsciiIgnoreCase(item.AnsiName, property.Value));
                if (module != null && !TryAppendDesignerStorage(compound, module.StreamName, buffer)) {
                    detail = "The VBA project-normalized designer transcript exceeds the configured byte limit.";
                    return false;
                }
            }
            if (IsExcludedProjectProperty(property.Name)) continue;
            if (!buffer.TryAppend(property.Name) || !buffer.TryAppend(property.Value)) {
                detail = "The VBA project-normalized transcript exceeds the configured byte limit.";
                return false;
            }
        }
        bool inHostExtender = false;
        foreach (byte[] raw in SplitLines(project)) {
            byte[] line = TrimAscii(raw);
            if (AsciiEquals(line, "[Host Extender Info]")) {
                inHostExtender = true;
                if (!buffer.TryAppend(Encoding.ASCII.GetBytes("Host Extender Info"))) {
                    detail = "The VBA host-extender transcript exceeds the configured byte limit.";
                    return false;
                }
                continue;
            }
            if (!inHostExtender) continue;
            if (line.Length > 1 && line[0] == (byte)'[' && line[line.Length - 1] == (byte)']') break;
            if (StartsWithAsciiIgnoreCase(line, "&H") && !buffer.TryAppend(line)) {
                detail = "The VBA host-extender transcript exceeds the configured byte limit.";
                return false;
            }
        }
        output = buffer.ToArray();
        detail = string.Empty;
        return true;
    }

    private static bool TryAppendDesignerStorage(OfficeCompoundFile compound, string storageName,
        BoundedBuffer buffer) {
        OfficeCompoundFileEntry? storage = compound.Entries.FirstOrDefault(entry => entry.IsStorage &&
            !entry.IsFallback && string.Equals(entry.Path, storageName, StringComparison.OrdinalIgnoreCase));
        if (storage == null) return true;
        string prefix = storage.Path + "/";
        foreach (OfficeCompoundFileEntry entry in compound.Entries.Where(entry => entry.IsStream &&
                     !entry.IsFallback && entry.Path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                     .OrderBy(entry => entry.DirectoryOrder)) {
            if (!compound.Streams.TryGetValue(entry.Path, out byte[]? bytes) || bytes.Length == 0) continue;
            if (!buffer.TryAppend(bytes)) return false;
            int padding = 1023 - bytes.Length % 1023;
            if (padding != 1023 && !buffer.TryAppend(new byte[padding])) return false;
        }
        return true;
    }

    private static bool TryReadModuleSource(OfficeCompoundFile compound, ModuleModel module,
        int maximumBytes, out byte[] source, out string detail) {
        source = Array.Empty<byte>();
        string path = "VBA/" + module.StreamName;
        if (!compound.Streams.TryGetValue(path, out byte[]? stream)) {
            detail = "The VBA module stream '" + path + "' is missing.";
            return false;
        }
        if (module.TextOffset > stream.Length) {
            detail = "The VBA module source offset is outside '" + path + "'.";
            return false;
        }
        var compressed = new byte[stream.Length - module.TextOffset];
        Buffer.BlockCopy(stream, module.TextOffset, compressed, 0, compressed.Length);
        return TryDecompress(compressed, maximumBytes, out source, out detail);
    }

    internal static bool TryDecompress(byte[] input, int maximumOutputBytes,
        out byte[] output, out string detail) {
        output = Array.Empty<byte>();
        if (input.Length == 0 || input[0] != 0x01) {
            detail = "The MS-OVBA compressed container signature is missing.";
            return false;
        }
        var decompressed = new List<byte>(Math.Min(input.Length * 2, maximumOutputBytes));
        int position = 1;
        while (position < input.Length) {
            int headerPosition = position;
            if (!TryReadUInt16(input, ref position, out ushort header)) {
                detail = "The compressed container ends inside a chunk header.";
                return false;
            }
            int chunkSize = (header & 0x0FFF) + 3;
            int chunkEnd = headerPosition + chunkSize;
            if ((header & 0x7000) != 0x3000 || chunkEnd < position || chunkEnd > input.Length) {
                detail = "The compressed container has an invalid chunk header.";
                return false;
            }
            int chunkOutputStart = decompressed.Count;
            if ((header & 0x8000) == 0) {
                if (chunkSize != 4098 || chunkEnd - position != 4096 ||
                    decompressed.Count > maximumOutputBytes - 4096) {
                    detail = "The compressed container has an invalid or oversized raw chunk.";
                    return false;
                }
                for (; position < chunkEnd; position++) decompressed.Add(input[position]);
                continue;
            }
            while (position < chunkEnd) {
                byte flags = input[position++];
                for (int bit = 0; bit < 8 && position < chunkEnd; bit++) {
                    if ((flags & 1 << bit) == 0) {
                        if (decompressed.Count >= maximumOutputBytes) {
                            detail = "The expanded MS-OVBA container exceeds the configured byte limit.";
                            return false;
                        }
                        decompressed.Add(input[position++]);
                        continue;
                    }
                    if (!TryReadUInt16(input, ref position, out ushort token) || position > chunkEnd) {
                        detail = "The compressed container ends inside a copy token.";
                        return false;
                    }
                    int decompressedPosition = decompressed.Count - chunkOutputStart;
                    int bitCount = 4;
                    while (bitCount < 12 && 1 << bitCount < decompressedPosition) bitCount++;
                    int lengthMask = 0xFFFF >> bitCount;
                    int offset = ((token & ~lengthMask) >> (16 - bitCount)) + 1;
                    int length = (token & lengthMask) + 3;
                    int sourceOffset = decompressed.Count - offset;
                    if (decompressedPosition <= 0 || sourceOffset < chunkOutputStart ||
                        decompressedPosition + length > 4096 || decompressed.Count > maximumOutputBytes - length) {
                        detail = "The compressed container has an out-of-range copy token.";
                        return false;
                    }
                    for (int copied = 0; copied < length; copied++) decompressed.Add(decompressed[sourceOffset + copied]);
                }
            }
        }
        output = decompressed.ToArray();
        detail = string.Empty;
        return true;
    }

    private sealed class BoundedBuffer {
        private readonly int _maximum;
        private readonly MemoryStream _stream = new();
        internal BoundedBuffer(int maximum) => _maximum = maximum;
        internal bool TryAppend(byte[] bytes) {
            if (bytes.Length > _maximum - _stream.Length) return false;
            _stream.Write(bytes, 0, bytes.Length);
            return true;
        }
        internal bool TryAppendByte(byte value) {
            if (_stream.Length >= _maximum) return false;
            _stream.WriteByte(value);
            return true;
        }
        internal byte[] ToArray() => _stream.ToArray();
    }

    private sealed class ProjectProperty {
        internal ProjectProperty(byte[] name, byte[] value) { Name = name; Value = value; }
        internal byte[] Name { get; }
        internal byte[] Value { get; }
    }

    private static IEnumerable<ProjectProperty> ReadProjectProperties(byte[] project) {
        foreach (byte[] raw in SplitLines(project)) {
            byte[] line = TrimAscii(raw);
            if (line.Length == 0) continue;
            if (line.Length >= 3 && line[0] == 0xEF && line[1] == 0xBB && line[2] == 0xBF) line = Slice(line, 3, line.Length - 3);
            if (line.Length > 1 && line[0] == (byte)'[' && line[line.Length - 1] == (byte)']') yield break;
            int equals = Array.IndexOf(line, (byte)'=');
            if (equals <= 0) continue;
            byte[] name = TrimAscii(Slice(line, 0, equals));
            byte[] value = TrimAscii(Slice(line, equals + 1, line.Length - equals - 1));
            if (value.Length >= 2 && value[0] == (byte)'"' && value[value.Length - 1] == (byte)'"') {
                value = Slice(value, 1, value.Length - 2);
            }
            yield return new ProjectProperty(name, value);
        }
    }

    private static IEnumerable<byte[]> SplitLines(byte[] bytes) {
        int start = 0;
        for (int index = 0; index < bytes.Length; index++) {
            if (bytes[index] != 0x0D && bytes[index] != 0x0A) continue;
            yield return Slice(bytes, start, index - start);
            if (index + 1 < bytes.Length && ((bytes[index] == 0x0D && bytes[index + 1] == 0x0A)
                || (bytes[index] == 0x0A && bytes[index + 1] == 0x0D))) index++;
            start = index + 1;
        }
        yield return Slice(bytes, start, bytes.Length - start);
    }

    private static bool IsExcludedProjectProperty(byte[] name) =>
        AsciiEquals(name, "ID") || AsciiEquals(name, "Document") || AsciiEquals(name, "DocModule")
        || AsciiEquals(name, "CMG") || AsciiEquals(name, "DPB") || AsciiEquals(name, "GC")
        || AsciiEquals(name, "ProtectionState") || AsciiEquals(name, "Password")
        || AsciiEquals(name, "VisibilityState");

    private static bool AsciiEquals(byte[] bytes, string value) =>
        bytes.Length == value.Length && StartsWithAsciiIgnoreCase(bytes, value);

    private static bool StartsWithAsciiIgnoreCase(byte[] bytes, string value) {
        if (bytes.Length < value.Length) return false;
        for (int index = 0; index < value.Length; index++) {
            byte left = bytes[index];
            byte right = (byte)value[index];
            if (left >= (byte)'A' && left <= (byte)'Z') left = (byte)(left + 32);
            if (right >= (byte)'A' && right <= (byte)'Z') right = (byte)(right + 32);
            if (left != right) return false;
        }
        return true;
    }

    private static bool BytesEqualAsciiIgnoreCase(byte[] left, byte[] right) {
        if (left.Length != right.Length) return false;
        for (int index = 0; index < left.Length; index++) {
            byte a = left[index];
            byte b = right[index];
            if (a >= (byte)'A' && a <= (byte)'Z') a = (byte)(a + 32);
            if (b >= (byte)'A' && b <= (byte)'Z') b = (byte)(b + 32);
            if (a != b) return false;
        }
        return true;
    }

    private static bool BytesEqual(byte[] left, byte[] right) => left.SequenceEqual(right);

    private static byte[] TrimAscii(byte[] bytes) {
        int start = 0;
        int end = bytes.Length;
        while (start < end && IsAsciiWhitespace(bytes[start])) start++;
        while (end > start && IsAsciiWhitespace(bytes[end - 1])) end--;
        return Slice(bytes, start, end - start);
    }

    private static bool IsAsciiWhitespace(byte value) => value is 0x09 or 0x0A or 0x0B or 0x0C or 0x0D or 0x20;

    private static byte[] CopyUntilNull(byte[] bytes) {
        int length = Array.IndexOf(bytes, (byte)0);
        return length < 0 ? bytes : Slice(bytes, 0, length);
    }

    private static bool TryReadUInt16(byte[] bytes, ref int position, out ushort value) {
        value = 0;
        if (position + 2 > bytes.Length) return false;
        value = ReadUInt16(bytes, position);
        position += 2;
        return true;
    }

    private static ushort ReadUInt16(byte[] bytes, int offset) =>
        (ushort)(bytes[offset] | bytes[offset + 1] << 8);

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        (uint)(bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24);

    private static byte[] UInt16Bytes(ushort value) => new[] { (byte)value, (byte)(value >> 8) };

    private static byte[] UInt32Bytes(uint value) => new[] {
        (byte)value, (byte)(value >> 8), (byte)(value >> 16), (byte)(value >> 24)
    };

    private static byte[] WidenBytes(byte[] bytes) {
        var widened = new byte[checked(bytes.Length * 2)];
        for (int index = 0; index < bytes.Length; index++) widened[index * 2] = bytes[index];
        return widened;
    }

    private static byte[] Slice(byte[] bytes, int offset, int count) {
        var result = new byte[count];
        if (count > 0) Buffer.BlockCopy(bytes, offset, result, 0, count);
        return result;
    }

    private static byte[] Concat(params byte[][] values) {
        int length = values.Aggregate(0, (current, value) => checked(current + value.Length));
        var output = new byte[length];
        int offset = 0;
        foreach (byte[] value in values) {
            Buffer.BlockCopy(value, 0, output, offset, value.Length);
            offset += value.Length;
        }
        return output;
    }
}
