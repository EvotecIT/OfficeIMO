using OfficeIMO.Excel.LegacyXls.Model;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffHyperlinkReader {
        private const uint StreamVersion = 2;
        private const uint HasMoniker = 0x00000001;
        private const uint HasLocation = 0x00000008;
        private const uint HasDisplayName = 0x00000010;
        private const uint HasGuid = 0x00000020;
        private const uint HasCreationTime = 0x00000040;
        private const uint HasFrameName = 0x00000080;
        private const uint MonikerSavedAsString = 0x00000100;

        private static readonly byte[] UrlMonikerClsid = {
            0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11,
            0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b
        };

        private static readonly byte[] FileMonikerClsid = {
            0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46
        };

        internal static bool TryRead(byte[] payload, out LegacyXlsHyperlink? hyperlink) {
            hyperlink = null;
            if (payload.Length < 32) {
                return false;
            }

            ushort firstRow = BiffRecordReader.ReadUInt16(payload, 0);
            ushort lastRow = BiffRecordReader.ReadUInt16(payload, 2);
            ushort firstColumn = BiffRecordReader.ReadUInt16(payload, 4);
            ushort lastColumn = BiffRecordReader.ReadUInt16(payload, 6);
            if (lastRow < firstRow || lastColumn < firstColumn) {
                return false;
            }

            int offset = 24;
            if (BiffRecordReader.ReadUInt32(payload, offset) != StreamVersion) {
                return false;
            }

            offset += 4;
            uint flags = BiffRecordReader.ReadUInt32(payload, offset);
            offset += 4;

            string? displayName = null;
            if ((flags & HasDisplayName) != 0 && !TryReadHyperlinkString(payload, ref offset, out displayName)) {
                return false;
            }

            if ((flags & HasFrameName) != 0 && !TrySkipHyperlinkString(payload, ref offset)) {
                return false;
            }

            string? target = null;
            if ((flags & HasMoniker) != 0) {
                if ((flags & MonikerSavedAsString) != 0) {
                    if (!TryReadHyperlinkString(payload, ref offset, out target)) {
                        return false;
                    }
                } else if (!TryReadMoniker(payload, ref offset, out target)) {
                    return false;
                }
            }

            string? location = null;
            if ((flags & HasLocation) != 0 && !TryReadHyperlinkString(payload, ref offset, out location)) {
                return false;
            }

            if ((flags & HasGuid) != 0) {
                if (offset + 16 > payload.Length) {
                    return false;
                }

                offset += 16;
            }

            if ((flags & HasCreationTime) != 0 && offset + 8 > payload.Length) {
                return false;
            }

            bool isExternal = !string.IsNullOrWhiteSpace(target);
            if (isExternal && !IsSupportedExternalTarget(target!)) {
                return false;
            }

            string hyperlinkTarget;
            if (isExternal) {
                hyperlinkTarget = target!;
                if (!string.IsNullOrWhiteSpace(location)) {
                    hyperlinkTarget += "#" + location;
                }
            } else if (!string.IsNullOrWhiteSpace(location)) {
                hyperlinkTarget = NormalizeInternalLocation(location!);
            } else {
                return false;
            }

            hyperlink = new LegacyXlsHyperlink(
                firstRow + 1,
                firstColumn + 1,
                lastRow + 1,
                lastColumn + 1,
                hyperlinkTarget,
                string.IsNullOrWhiteSpace(displayName) ? null : displayName,
                isExternal);
            return true;
        }

        private static string NormalizeInternalLocation(string location) {
            string normalized = location.Trim();
            return normalized.StartsWith("#", StringComparison.Ordinal)
                ? normalized.Substring(1)
                : normalized;
        }

        private static bool TryReadHyperlinkString(byte[] payload, ref int offset, out string? value) {
            value = null;
            if (offset + 4 > payload.Length) {
                return false;
            }

            uint characterCount = BiffRecordReader.ReadUInt32(payload, offset);
            offset += 4;
            if (characterCount == 0 || characterCount > int.MaxValue / 2) {
                return false;
            }

            int byteCount = checked((int)characterCount * 2);
            if (offset + byteCount > payload.Length) {
                return false;
            }

            value = Encoding.Unicode.GetString(payload, offset, byteCount).TrimEnd('\0');
            offset += byteCount;
            return true;
        }

        private static bool TrySkipHyperlinkString(byte[] payload, ref int offset) {
            return TryReadHyperlinkString(payload, ref offset, out _);
        }

        private static bool TryReadMoniker(byte[] payload, ref int offset, out string? target) {
            if (offset + 16 > payload.Length) {
                target = null;
                return false;
            }

            if (HasClsid(payload, offset, UrlMonikerClsid)) {
                return TryReadUrlMoniker(payload, ref offset, out target);
            }

            if (HasClsid(payload, offset, FileMonikerClsid)) {
                return TryReadFileMoniker(payload, ref offset, out target);
            }

            target = null;
            return false;
        }

        private static bool TryReadUrlMoniker(byte[] payload, ref int offset, out string? target) {
            target = null;
            if (offset + 20 > payload.Length || !HasClsid(payload, offset, UrlMonikerClsid)) {
                return false;
            }

            offset += 16;
            uint byteCount = BiffRecordReader.ReadUInt32(payload, offset);
            offset += 4;
            if (byteCount == 0 || byteCount > int.MaxValue || offset + byteCount > payload.Length) {
                return false;
            }

            int stringByteCount = (int)byteCount;
            if (stringByteCount >= 24) {
                uint possibleStringByteCount = byteCount - 24;
                if (possibleStringByteCount > 0 && possibleStringByteCount % 2 == 0 && HasTerminatingNull(payload, offset, (int)possibleStringByteCount)) {
                    stringByteCount = (int)possibleStringByteCount;
                }
            }

            if (stringByteCount % 2 != 0 || !HasTerminatingNull(payload, offset, stringByteCount)) {
                return false;
            }

            target = Encoding.Unicode.GetString(payload, offset, stringByteCount).TrimEnd('\0');
            offset += (int)byteCount;
            return true;
        }

        private static bool TryReadFileMoniker(byte[] payload, ref int offset, out string? target) {
            target = null;
            if (offset + 22 > payload.Length || !HasClsid(payload, offset, FileMonikerClsid)) {
                return false;
            }

            offset += 16;
            offset += 2; // cAnti: parent-directory indicator count for relative file monikers.
            uint ansiLength = BiffRecordReader.ReadUInt32(payload, offset);
            offset += 4;
            if (ansiLength == 0 || ansiLength > 32767 || offset + ansiLength > payload.Length) {
                return false;
            }

            string ansiPath = Encoding.ASCII.GetString(payload, offset, (int)ansiLength).TrimEnd('\0');
            offset += (int)ansiLength;
            if (offset + 28 > payload.Length) {
                return false;
            }

            offset += 2;
            ushort version = BiffRecordReader.ReadUInt16(payload, offset);
            offset += 2;
            if (version != 0xdead) {
                return false;
            }

            offset += 20;
            uint unicodePathSize = BiffRecordReader.ReadUInt32(payload, offset);
            offset += 4;

            if (unicodePathSize > 0) {
                if (unicodePathSize < 6 || offset + unicodePathSize > payload.Length) {
                    return false;
                }

                uint unicodePathBytes = BiffRecordReader.ReadUInt32(payload, offset);
                offset += 4;
                ushort key = BiffRecordReader.ReadUInt16(payload, offset);
                offset += 2;
                if (key != 3 || unicodePathBytes == 0 || unicodePathBytes % 2 != 0 || offset + unicodePathBytes > payload.Length) {
                    return false;
                }

                target = Encoding.Unicode.GetString(payload, offset, (int)unicodePathBytes);
                offset += (int)unicodePathBytes;
            } else {
                target = ansiPath;
            }

            target = target.TrimEnd('\0');
            return !string.IsNullOrWhiteSpace(target);
        }

        private static bool HasClsid(byte[] payload, int offset, byte[] expected) {
            for (int i = 0; i < expected.Length; i++) {
                if (payload[offset + i] != expected[i]) {
                    return false;
                }
            }

            return true;
        }

        private static bool HasTerminatingNull(byte[] payload, int offset, int byteCount) {
            return byteCount >= 2 && payload[offset + byteCount - 1] == 0 && payload[offset + byteCount - 2] == 0;
        }

        private static bool IsSupportedExternalTarget(string target) {
            string trimmed = target.Trim();
            return IsSupportedAbsoluteExternalUri(trimmed) || IsSupportedRelativeFileTarget(trimmed);
        }

        private static bool IsSupportedAbsoluteExternalUri(string target) {
            if (!Uri.TryCreate(target, UriKind.Absolute, out Uri? uri)) {
                return false;
            }

            return string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                || string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)
                || string.Equals(uri.Scheme, Uri.UriSchemeFtp, StringComparison.OrdinalIgnoreCase)
                || string.Equals(uri.Scheme, Uri.UriSchemeMailto, StringComparison.OrdinalIgnoreCase)
                || string.Equals(uri.Scheme, Uri.UriSchemeFile, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsSupportedRelativeFileTarget(string target) {
            if (string.IsNullOrWhiteSpace(target) || target.StartsWith("#", StringComparison.Ordinal)) {
                return false;
            }

            if (target.StartsWith("/", StringComparison.Ordinal) || target.StartsWith("\\", StringComparison.Ordinal)) {
                return false;
            }

            int colonIndex = target.IndexOf(':');
            if (colonIndex >= 0) {
                int slashIndex = target.IndexOfAny(new[] { '/', '\\' });
                if (slashIndex < 0 || colonIndex < slashIndex) {
                    return false;
                }
            }

            return Uri.TryCreate(target.Replace('\\', '/'), UriKind.Relative, out _);
        }
    }
}
