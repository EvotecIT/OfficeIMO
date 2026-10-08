using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeIMO.Core.Internal;

/// <summary>One directory decoder for VBA source editing and signature-binding transcripts.</summary>
internal static class OfficeVbaDirectoryCodec {
    internal sealed class DirectoryModel {
        internal byte[] SerializedPrefix = Array.Empty<byte>();
        internal int CodePage;
        internal byte[] V3Prefix = Array.Empty<byte>();
        internal byte[] ProjectModulesHeader = Array.Empty<byte>();
        internal byte[] ProjectName = Array.Empty<byte>();
        internal byte[] ProjectConstants = Array.Empty<byte>();
        internal byte[] ProjectCookieHeader = Array.Empty<byte>();
        internal byte[] SerializedCookie = Array.Empty<byte>();
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
                || !v3.TryAppend(sysKindHeader)) {
                detail = "The VBA directory has invalid project system or locale records.";
                return false;
            }
            // Modern Office inserts PROJECTCOMPATVERSION before the locale records.
            if (reader.PeekId == 0x004a && (!reader.TryReadSized(0x004a, out byte[] compatibility, out _, includeData: true) || compatibility.Length != 4)) {
                detail = "The VBA directory has an invalid compatibility-version record.";
                return false;
            }
            if (!reader.TryReadSized(0x0002, out _, out byte[] lcidRecord, includeData: true) || !v3.TryAppend(lcidRecord)) {
                detail = "The VBA directory has an invalid locale record.";
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
                || !reader.TryReadSized(0x0004, out parsed.ProjectName, out byte[] projectNameRecord, true)
                || !reader.TryReadSized(0x0005, out _, out byte[] docStringHeader, false)
                || !reader.TryReadSized(0x0040, out _, out byte[] docStringUnicodeHeader, false)
                || !reader.TryReadSized(0x0006, out _, out byte[] helpFileHeader, false)
                || !reader.TryReadSized(0x003D, out _, out byte[] helpFileUnicodeHeader, false)
                || !reader.TryReadSized(0x0007, out _, out byte[] helpContextHeader, false)
                || !reader.TryReadSized(0x0008, out _, out byte[] libFlagsRecord, true)
                || !reader.TryReadProjectVersion(out byte[] versionRecord)
                || !v3.TryAppend(codePageHeader) || !v3.TryAppend(projectNameRecord)
                || !v3.TryAppend(docStringHeader) || !v3.TryAppend(docStringUnicodeHeader)
                || !v3.TryAppend(helpFileHeader) || !v3.TryAppend(helpFileUnicodeHeader)
                || !v3.TryAppend(helpContextHeader) || !v3.TryAppend(libFlagsRecord)
                || !v3.TryAppend(versionRecord)) {
                detail = "The VBA directory has invalid project metadata records.";
                return false;
            }
            if (codePage.Length != 2) {
                detail = "The VBA project code page record must contain two bytes.";
                return false;
            }
            parsed.CodePage = ReadUInt16(codePage, 0);
            reader.CodePage = parsed.CodePage;
            if (reader.PeekId == 0x000c && (!reader.TryReadSized(0x000c, out parsed.ProjectConstants, out byte[] constantsRecord, true)
                || !reader.TryReadSized(0x003c, out _, out byte[] constantsUnicodeRecord, true)
                || !v3.TryAppend(constantsRecord) || !v3.TryAppend(constantsUnicodeRecord))) {
                detail = "The VBA directory has invalid conditional-compilation constants.";
                return false;
            }
            parsed.SerializedPrefix = reader.GetBytes(0, reader.Position);
            while (reader.PeekId != 0x000F) {
                int referenceStart = reader.Position;
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
                reference.Serialized = reader.GetBytes(referenceStart, reader.Position - referenceStart);
                parsed.References.Add(reference);
                if (!v3.TryAppend(reference.V3Normalized)) {
                    detail = "The VBA V3 reference transcript exceeds the configured byte limit.";
                    return false;
                }
            }
            if (!reader.TryReadSized(0x000F, out byte[] moduleCountBytes, out parsed.ProjectModulesHeader, false)
                || moduleCountBytes.Length != 2
                || !reader.TryReadSized(0x0013, out byte[] cookie, out parsed.ProjectCookieHeader, false) || cookie.Length != 2) {
                detail = "The VBA directory has invalid module-count or project-cookie records.";
                return false;
            }
            parsed.SerializedCookie = Concat(parsed.ProjectCookieHeader, cookie);
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
        internal byte[] Serialized = Array.Empty<byte>();
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
        internal byte[] Serialized = Array.Empty<byte>();
        internal byte[] LegacyNormalized = Array.Empty<byte>();
        internal byte[] V3Normalized = Array.Empty<byte>();
        internal byte[] Normalize() => LegacyNormalized;
    }

    private sealed class DirectoryReader {
        private readonly byte[] _bytes;
        private int _position;
        internal DirectoryReader(byte[] bytes) => _bytes = bytes;
        internal int CodePage { get; set; } = 1252;
        internal bool AtEnd {
            get {
                // The standard raw final compression chunk may contain zero padding after the terminator.
                for (int index = _position; index < _bytes.Length; index++) if (_bytes[index] != 0) return false;
                return true;
            }
        }
        internal int Position => _position;
        internal byte[] GetBytes(int offset, int count) => Slice(_bytes, offset, count);
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
            int start = _position;
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
                    ? new UnicodeEncoding(false, false, true).GetString(unicodeStream).TrimEnd('\0')
                    : OfficeVbaText.Decode(ansiStream, CodePage).TrimEnd('\0');
            } catch (DecoderFallbackException) { return false; }
            catch (NotSupportedException) { return false; }
            uint textOffset = ReadUInt32(textOffsetBytes, 0);
            if (textOffset > int.MaxValue) return false;
            model = new ModuleModel {
                Serialized = Slice(_bytes, start, _position - start),
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

    private static byte[] Slice(byte[] bytes, int offset, int count) { var result = new byte[count]; Buffer.BlockCopy(bytes, offset, result, 0, count); return result; }
    private static ushort ReadUInt16(byte[] bytes, int offset) => (ushort)(bytes[offset] | bytes[offset + 1] << 8);
    private static uint ReadUInt32(byte[] bytes, int offset) => (uint)(bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24);
    private static byte[] UInt16Bytes(ushort value) => new byte[] { (byte)value, (byte)(value >> 8) };
    private static byte[] UInt32Bytes(uint value) => new byte[] { (byte)value, (byte)(value >> 8), (byte)(value >> 16), (byte)(value >> 24) };
    private static byte[] Concat(params byte[][] values) { using var result = new MemoryStream(); foreach (byte[] value in values) result.Write(value, 0, value.Length); return result.ToArray(); }
    private static byte[] WidenBytes(byte[] bytes) { var result = new byte[bytes.Length * 2]; for (int index = 0; index < bytes.Length; index++) result[index * 2] = bytes[index]; return result; }
    private static byte[] CopyUntilNull(byte[] bytes) { int count = Array.IndexOf(bytes, (byte)0); return count < 0 ? bytes : Slice(bytes, 0, count); }
    private static bool TryReadUInt16(byte[] bytes, ref int position, out ushort value) { value = 0; if (position < 0 || position > bytes.Length - 2) return false; value = ReadUInt16(bytes, position); position += 2; return true; }
}
