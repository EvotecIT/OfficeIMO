using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Security;

/// <summary>Bounded managed implementation of the MS-OVBA signature-binding transcripts.</summary>
internal static class OfficeVbaProjectCanonicalizer {
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
            byte[][] lines = SplitLines(source).ToArray();
            for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++) {
                byte[] line = lines[lineIndex];
                if (lineIndex == lines.Length - 1 && line.Length == 0) continue;
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

    private sealed class DirectoryModel {
        internal byte[] V3Prefix = Array.Empty<byte>();
        internal byte[] ProjectModulesHeader = Array.Empty<byte>();
        internal byte[] ProjectName = Array.Empty<byte>();
        internal byte[] ProjectConstants = Array.Empty<byte>();
        internal byte[] ProjectCookieHeader = Array.Empty<byte>();
        internal byte[] TerminatorRecord = Array.Empty<byte>();
        internal readonly List<ReferenceModel> References = new();
        internal readonly List<ModuleModel> Modules = new();

        internal static bool TryParse(byte[] bytes, int maximumBytes,
            out DirectoryModel? model, out string detail) {
            model = null;
            var reader = new DirectoryReader(bytes);
            var parsed = new DirectoryModel();
            var v3 = new BoundedBuffer(maximumBytes);
            if (!reader.TryReadSized(0x0001, out _, out byte[] sysKindHeader, includeData: false)
                || !reader.TryReadSized(0x0002, out _, out byte[] lcidRecord, includeData: true)
                || !v3.TryAppend(sysKindHeader) || !v3.TryAppend(lcidRecord)) {
                detail = "The VBA directory has invalid project system or locale records.";
                return false;
            }
            if (reader.PeekId == 0x0014) {
                if (!reader.TryReadSized(0x0014, out _, out byte[] invokeRecord, true)
                    || !v3.TryAppend(invokeRecord)) {
                    detail = "The VBA directory has an invalid invoke-locale record.";
                    return false;
                }
            }
            if (!reader.TryReadSized(0x0003, out _, out byte[] codePageHeader, false)
                || !reader.TryReadSized(0x0004, out parsed.ProjectName, out byte[] projectNameRecord, true)
                || !reader.TryReadSized(0x0005, out _, out byte[] docStringHeader, false)
                || !reader.TryReadSized(0x0040, out _, out byte[] docStringUnicodeHeader, false)
                || !reader.TryReadSized(0x0006, out _, out byte[] helpFileHeader, false)
                || !reader.TryReadSized(0x003D, out _, out byte[] helpFileUnicodeHeader, false)
                || !reader.TryReadSized(0x0007, out _, out byte[] helpContextHeader, false)
                || !reader.TryReadSized(0x0008, out _, out byte[] libFlagsRecord, true)
                || !reader.TryReadProjectVersion(out byte[] versionRecord)
                || !reader.TryReadSized(0x000C, out parsed.ProjectConstants, out byte[] constantsRecord, true)
                || !reader.TryReadSized(0x003C, out _, out byte[] constantsUnicodeRecord, true)
                || !v3.TryAppend(codePageHeader) || !v3.TryAppend(projectNameRecord)
                || !v3.TryAppend(docStringHeader) || !v3.TryAppend(docStringUnicodeHeader)
                || !v3.TryAppend(helpFileHeader) || !v3.TryAppend(helpFileUnicodeHeader)
                || !v3.TryAppend(helpContextHeader) || !v3.TryAppend(libFlagsRecord)
                || !v3.TryAppend(versionRecord) || !v3.TryAppend(constantsRecord)
                || !v3.TryAppend(constantsUnicodeRecord)) {
                detail = "The VBA directory has invalid project metadata records.";
                return false;
            }
            while (reader.PeekId != 0x000F) {
                byte[] nameRecord = Array.Empty<byte>();
                if (reader.PeekId == 0x0016) {
                    if (!reader.TryReadSized(0x0016, out _, out byte[] nameAnsi, true)
                        || !reader.TryReadSized(0x003E, out _, out byte[] nameUnicode, true)) {
                        detail = "The VBA directory has an invalid reference-name record.";
                        return false;
                    }
                    nameRecord = Concat(nameAnsi, nameUnicode);
                }
                if (!reader.TryReadReference(nameRecord, out ReferenceModel? reference) || reference == null) {
                    detail = "The VBA directory has an invalid or unsupported reference record near 0x" +
                        reader.PeekId.ToString("X4", System.Globalization.CultureInfo.InvariantCulture) +
                        " at byte " + reader.Position.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".";
                    return false;
                }
                parsed.References.Add(reference);
                if (!v3.TryAppend(reference.V3Normalized)) {
                    detail = "The VBA V3 reference transcript exceeds the configured byte limit.";
                    return false;
                }
            }
            if (!reader.TryReadSized(0x000F, out byte[] moduleCountBytes, out parsed.ProjectModulesHeader, false)
                || moduleCountBytes.Length != 2
                || !reader.TryReadSized(0x0013, out _, out parsed.ProjectCookieHeader, false)) {
                detail = "The VBA directory has invalid module-count or project-cookie records.";
                return false;
            }
            parsed.V3Prefix = v3.ToArray();
            int moduleCount = moduleCountBytes[0] | moduleCountBytes[1] << 8;
            if (moduleCount > 4096) {
                detail = "The VBA directory exceeds the supported module count.";
                return false;
            }
            for (int index = 0; index < moduleCount; index++) {
                if (!reader.TryReadModule(out ModuleModel? module) || module == null) {
                    detail = "The VBA directory has an invalid module record at index " + index + ".";
                    return false;
                }
                parsed.Modules.Add(module);
            }
            if (!reader.TryReadFixedRecord(0x0010, out parsed.TerminatorRecord) || !reader.AtEnd) {
                detail = "The VBA directory has an invalid terminator or trailing data.";
                return false;
            }
            model = parsed;
            detail = string.Empty;
            return true;
        }
    }

    private sealed class ModuleModel {
        internal byte[] AnsiName = Array.Empty<byte>();
        internal byte[] UnicodeName = Array.Empty<byte>();
        internal string StreamName = string.Empty;
        internal int TextOffset;
        internal ushort TypeId;
        internal byte[] TypeRecord = Array.Empty<byte>();
        internal byte[]? ReadOnlyRecord;
        internal byte[]? PrivateRecord;
        internal byte[] PreferredNameBytes => UnicodeName.Length > 0 ? UnicodeName : AnsiName;
    }

    private sealed class ReferenceModel {
        internal byte[] LegacyNormalized = Array.Empty<byte>();
        internal byte[] V3Normalized = Array.Empty<byte>();
        internal byte[] Normalize() => LegacyNormalized;
    }

    private sealed class DirectoryReader {
        private readonly byte[] _bytes;
        private int _position;
        internal DirectoryReader(byte[] bytes) => _bytes = bytes;
        internal bool AtEnd => _position == _bytes.Length;
        internal int Position => _position;
        internal ushort PeekId => _position + 2 <= _bytes.Length ? ReadUInt16(_bytes, _position) : ushort.MaxValue;

        internal bool TryReadSized(ushort expectedId, out byte[] data, out byte[] header, bool includeData) {
            data = Array.Empty<byte>();
            header = Array.Empty<byte>();
            int start = _position;
            if (!TryReadU16(out ushort id) || id != expectedId || !TryReadU32(out uint size)
                || size > int.MaxValue || _position + (long)size > _bytes.Length) return false;
            data = Slice(_bytes, _position, (int)size);
            _position += (int)size;
            header = Slice(_bytes, start, includeData ? _position - start : 6);
            return true;
        }

        internal bool TryReadProjectVersion(out byte[] record) {
            record = Array.Empty<byte>();
            int start = _position;
            if (!TryReadU16(out ushort id) || id != 0x0009 || !TryReadU32(out uint reserved)
                || reserved != 4 || !TryReadU32(out _) || !TryReadU16(out _)) return false;
            record = Slice(_bytes, start, _position - start);
            return true;
        }

        internal bool TryReadReference(byte[] nameRecord, out ReferenceModel? model) {
            model = null;
            if (PeekId == 0x0033) {
                if (!TryReadU16(out ushort id) || id != 0x0033
                    || !TryReadLengthPrefixed(out byte[] original, out byte[] originalLength)
                    || PeekId != 0x002F
                    || !TryReadReference(Array.Empty<byte>(), out ReferenceModel? nestedControl)
                    || nestedControl == null) return false;
                model = new ReferenceModel {
                    V3Normalized = Concat(nameRecord, UInt16Bytes(id), originalLength, original,
                        nestedControl.V3Normalized)
                };
                return true;
            }
            if (PeekId == 0x000D) {
                if (!TryReadU16(out ushort id) || id != 0x000D
                    || !TryReadU32(out _)
                    || !TryReadLengthPrefixed(out byte[] libid, out byte[] libidLength)
                    || !TryReadU32(out uint reserved1) || reserved1 != 0
                    || !TryReadU16(out ushort reserved2) || reserved2 != 0) return false;
                model = new ReferenceModel {
                    LegacyNormalized = new byte[] { 0x7B },
                    V3Normalized = Concat(nameRecord, UInt16Bytes(id), libidLength, WidenBytes(libid),
                        UInt32Bytes(reserved1), UInt16Bytes(reserved2))
                };
                return true;
            }
            if (PeekId == 0x000E) {
                if (!TryReadU16(out ushort id) || id != 0x000E || !TryReadU32(out _)
                    || !TryReadLengthPrefixed(out byte[] absolute, out byte[] absoluteLength)
                    || !TryReadLengthPrefixed(out byte[] relative, out byte[] relativeLength)
                    || !TryReadU32(out uint major) || !TryReadU16(out ushort minor)) return false;
                byte[] body = Concat(absoluteLength, absolute, relativeLength, relative,
                    UInt32Bytes(major), UInt16Bytes(minor));
                model = new ReferenceModel {
                    LegacyNormalized = CopyUntilNull(body),
                    V3Normalized = Concat(nameRecord, UInt16Bytes(id), body)
                };
                return true;
            }
            if (PeekId == 0x002F) {
                if (!TryReadU16(out ushort id) || id != 0x002F || !TryReadU32(out _)
                    || !TryReadLengthPrefixed(out byte[] twiddled, out byte[] twiddledLength)
                    || !TryReadU32(out uint reserved1) || reserved1 != 0
                    || !TryReadU16(out ushort reserved2) || reserved2 != 0) return false;
                byte[] extendedName = Array.Empty<byte>();
                if (PeekId == 0x0016) {
                    if (!TryReadSized(0x0016, out _, out byte[] extendedNameAnsi, true)
                        || !TryReadSized(0x003E, out _, out byte[] extendedNameUnicode, true)) return false;
                    extendedName = Concat(extendedNameAnsi, extendedNameUnicode);
                }
                if (!TryReadU16(out ushort extendedId) || extendedId != 0x0030
                    || !TryReadU32(out _)
                    || !TryReadLengthPrefixed(out byte[] extended, out byte[] extendedLength)
                    || !TryReadU32(out uint reserved4) || reserved4 != 0
                    || !TryReadU16(out ushort reserved5) || reserved5 != 0
                    || !TryReadBytes(20, out byte[] typeLibAndCookie)) return false;
                model = new ReferenceModel {
                    V3Normalized = Concat(nameRecord, UInt16Bytes(id), twiddledLength, twiddled,
                        UInt32Bytes(reserved1), UInt16Bytes(reserved2), extendedName,
                        UInt16Bytes(extendedId), extendedLength, extended,
                        UInt32Bytes(reserved4), UInt16Bytes(reserved5), typeLibAndCookie)
                };
                return true;
            }
            return false;
        }

        internal bool TryReadModule(out ModuleModel? model) {
            model = null;
            if (!TryReadSized(0x0019, out byte[] ansiName, out _, false)) return false;
            byte[] unicodeName = Array.Empty<byte>();
            if (PeekId == 0x0047 && !TryReadSized(0x0047, out unicodeName, out _, false)) return false;
            if (!TryReadSized(0x001A, out byte[] ansiStream, out _, false)
                || !TryReadSized(0x0032, out byte[] unicodeStream, out _, false)
                || !TryReadSized(0x001C, out _, out _, false)
                || !TryReadSized(0x0048, out _, out _, false)
                || !TryReadSized(0x0031, out byte[] textOffsetBytes, out _, false) || textOffsetBytes.Length != 4
                || !TryReadSized(0x001E, out _, out _, false)
                || !TryReadSized(0x002C, out _, out _, false)) return false;
            ushort typeId = PeekId;
            if (typeId != 0x0021 && typeId != 0x0022 || !TryReadFixedRecord(typeId, out byte[] typeRecord)) return false;
            byte[]? readOnly = null;
            byte[]? privateRecord = null;
            if (PeekId == 0x0025 && !TryReadFixedRecord(0x0025, out readOnly)) return false;
            if (PeekId == 0x0028 && !TryReadFixedRecord(0x0028, out privateRecord)) return false;
            if (!TryReadFixedRecord(0x002B, out _)) return false;
            string streamName;
            try {
                streamName = unicodeStream.Length > 0
                    ? Encoding.Unicode.GetString(unicodeStream).TrimEnd('\0')
                    : Encoding.ASCII.GetString(ansiStream).TrimEnd('\0');
            } catch (DecoderFallbackException) { return false; }
            uint textOffset = ReadUInt32(textOffsetBytes, 0);
            if (textOffset > int.MaxValue) return false;
            model = new ModuleModel {
                AnsiName = ansiName,
                UnicodeName = unicodeName,
                StreamName = streamName,
                TextOffset = (int)textOffset,
                TypeId = typeId,
                TypeRecord = typeRecord,
                ReadOnlyRecord = readOnly,
                PrivateRecord = privateRecord
            };
            return streamName.Length > 0;
        }

        internal bool TryReadFixedRecord(ushort expectedId, out byte[] record) {
            record = Array.Empty<byte>();
            int start = _position;
            if (!TryReadU16(out ushort id) || id != expectedId || !TryReadU32(out _)) return false;
            record = Slice(_bytes, start, 6);
            return true;
        }

        private bool TryReadLengthPrefixed(out byte[] data, out byte[] lengthBytes) {
            data = Array.Empty<byte>();
            lengthBytes = Array.Empty<byte>();
            int start = _position;
            if (!TryReadU32(out uint size) || size > int.MaxValue || _position + (long)size > _bytes.Length) return false;
            lengthBytes = Slice(_bytes, start, 4);
            data = Slice(_bytes, _position, (int)size);
            _position += (int)size;
            return true;
        }

        private bool TryReadU16(out ushort value) {
            value = 0;
            if (_position + 2 > _bytes.Length) return false;
            value = ReadUInt16(_bytes, _position);
            _position += 2;
            return true;
        }

        private bool TryReadU32(out uint value) {
            value = 0;
            if (_position + 4 > _bytes.Length) return false;
            value = ReadUInt32(_bytes, _position);
            _position += 4;
            return true;
        }

        private bool TryReadBytes(int count, out byte[] bytes) {
            bytes = Array.Empty<byte>();
            if (count < 0 || _position + (long)count > _bytes.Length) return false;
            bytes = Slice(_bytes, _position, count);
            _position += count;
            return true;
        }
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
