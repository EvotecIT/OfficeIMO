using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Security;

internal static partial class OfficeVbaProjectCanonicalizer {
    internal sealed class DirectoryModel {
        internal ushort CodePage;
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
            if (!reader.TryReadSized(0x0003, out byte[] codePage, out byte[] codePageHeader, false)
                || codePage.Length != 2
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
            parsed.CodePage = ReadUInt16(codePage, 0);
            while (reader.PeekId != 0x000F) {
                byte[] nameRecord = Array.Empty<byte>();
                byte[] referenceName = Array.Empty<byte>();
                if (reader.PeekId == 0x0016) {
                    if (!reader.TryReadSized(0x0016, out _, out byte[] nameAnsi, true)
                        || !reader.TryReadSized(0x003E, out referenceName, out byte[] nameUnicode, true)) {
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
                reference.UnicodeName = referenceName;
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

    internal sealed class ModuleModel {
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

    internal sealed class ReferenceModel {
        internal byte[] UnicodeName = Array.Empty<byte>();
        internal byte[] LibId = Array.Empty<byte>();
        internal ushort Kind;
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
                    Kind = id, LibId = original,
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
                    Kind = id, LibId = libid,
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
                    Kind = id, LibId = absolute,
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
                    Kind = id, LibId = extended,
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

}
