using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Biff;
using DocumentFormat.OpenXml;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsHyperlinkWriter {
        private const int Biff8MaxRow = 65536;
        private const int Biff8MaxColumn = 256;
        private const uint StreamVersion = 2;
        private const uint HasMoniker = 0x00000001;
        private const uint IsAbsolute = 0x00000002;
        private const uint HasLocation = 0x00000008;
        private const uint HasDisplayName = 0x00000010;
        private const int MaxTooltipCharacters = 255;

        private static readonly byte[] HLinkClsid = {
            0xd0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11,
            0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b
        };

        private static readonly byte[] UrlMonikerClsid = {
            0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11,
            0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b
        };

        private static readonly byte[] FileMonikerClsid = {
            0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46
        };

        internal static IReadOnlyList<LegacyXlsHyperlinkRecord> CreateHyperlinkRecords(ExcelSheet sheet) {
            if (!TryCreateHyperlinkRecords(sheet, out IReadOnlyList<LegacyXlsHyperlinkRecord> records, out string? reason)) {
                throw new NotSupportedException($"Native XLS saving does not yet support {reason} on worksheet '{sheet.Name}'. Save as .xlsx or remove this feature before saving as .xls.");
            }

            return records;
        }

        internal static bool SupportsWorksheetHyperlinks(ExcelSheet sheet, out string? reason) {
            return TryCreateHyperlinkRecords(sheet, out _, out reason);
        }

        private static bool TryCreateHyperlinkRecords(ExcelSheet sheet, out IReadOnlyList<LegacyXlsHyperlinkRecord> payloads, out string? reason) {
            payloads = Array.Empty<LegacyXlsHyperlinkRecord>();
            reason = null;

            WorksheetPart worksheetPart = sheet.WorksheetPart;
            Worksheet? worksheet = worksheetPart.Worksheet;
            Hyperlinks? hyperlinks = worksheet?.Elements<Hyperlinks>().FirstOrDefault();
            if (hyperlinks == null) {
                return true;
            }

            if (!SupportsHyperlinksCollection(hyperlinks)) {
                reason = "hyperlink collection metadata";
                return false;
            }

            Dictionary<string, HyperlinkRelationship> relationships = worksheetPart.HyperlinkRelationships
                .Where(relationship => !string.IsNullOrWhiteSpace(relationship.Id))
                .ToDictionary(relationship => relationship.Id!, relationship => relationship, StringComparer.OrdinalIgnoreCase);

            var records = new List<LegacyXlsHyperlinkRecord>();
            foreach (Hyperlink hyperlink in hyperlinks.Elements<Hyperlink>()) {
                if (!SupportsHyperlinkMetadata(hyperlink)) {
                    reason = "hyperlink metadata";
                    return false;
                }

                string? reference = hyperlink.Reference?.Value;
                if (!TryParseBiff8Reference(reference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    reason = "hyperlink references outside the BIFF8 worksheet grid";
                    return false;
                }

                string? displayName = string.IsNullOrWhiteSpace(hyperlink.Display?.Value) ? null : hyperlink.Display!.Value;
                string? tooltip = string.IsNullOrWhiteSpace(hyperlink.Tooltip?.Value) ? null : hyperlink.Tooltip!.Value;
                string? location = string.IsNullOrWhiteSpace(hyperlink.Location?.Value) ? null : hyperlink.Location!.Value;
                string? relationshipId = hyperlink.Id?.Value;
                if (!string.IsNullOrWhiteSpace(relationshipId)) {
                    if (!relationships.TryGetValue(relationshipId!, out HyperlinkRelationship? relationship)
                        || relationship.Uri == null
                        || !relationship.IsExternal) {
                        reason = "hyperlink relationship targets";
                        return false;
                    }

                    string target = relationship.Uri.OriginalString;
                    if (IsSupportedAbsoluteExternalUri(target)) {
                        if (!TryAddHyperlinkRecords(records, firstRow, firstColumn, lastRow, lastColumn, BuildExternalUrlPayload(firstRow, firstColumn, lastRow, lastColumn, target, displayName, location), tooltip, out reason)) {
                            return false;
                        }

                        continue;
                    }

                    if (IsSupportedRelativeFileTarget(target)) {
                        if (!TryAddHyperlinkRecords(records, firstRow, firstColumn, lastRow, lastColumn, BuildFileMonikerPayload(firstRow, firstColumn, lastRow, lastColumn, target, displayName, location), tooltip, out reason)) {
                            return false;
                        }

                        continue;
                    }

                    {
                        reason = "unsupported external hyperlink targets";
                        return false;
                    }
                }

                if (!string.IsNullOrWhiteSpace(location)) {
                    if (!TryAddHyperlinkRecords(records, firstRow, firstColumn, lastRow, lastColumn, BuildInternalLocationPayload(firstRow, firstColumn, lastRow, lastColumn, NormalizeInternalLocation(location!), displayName), tooltip, out reason)) {
                        return false;
                    }

                    continue;
                }

                reason = "hyperlinks without a target";
                return false;
            }

            payloads = records;
            return true;
        }

        private static bool SupportsHyperlinksCollection(Hyperlinks hyperlinks) {
            if (hyperlinks.ChildElements.Any(child => child is not Hyperlink)) {
                return false;
            }

            foreach (OpenXmlAttribute attribute in hyperlinks.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                if (!string.Equals(attribute.LocalName, "count", StringComparison.Ordinal)) {
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsHyperlinkMetadata(Hyperlink hyperlink) {
            if (hyperlink.HasChildren) {
                return false;
            }

            foreach (OpenXmlAttribute attribute in hyperlink.GetAttributes()) {
                if (string.Equals(attribute.LocalName, "id", StringComparison.Ordinal)) {
                    if (!string.Equals(attribute.NamespaceUri, "http://schemas.openxmlformats.org/officeDocument/2006/relationships", StringComparison.Ordinal)) {
                        return false;
                    }

                    continue;
                }

                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                switch (attribute.LocalName) {
                    case "ref":
                    case "location":
                    case "tooltip":
                    case "display":
                        break;
                    default:
                        return false;
                }
            }

            return true;
        }

        private static bool TryAddHyperlinkRecords(
            List<LegacyXlsHyperlinkRecord> records,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            byte[] hyperlinkPayload,
            string? tooltip,
            out string? reason) {
            reason = null;
            if (hyperlinkPayload.Length > ushort.MaxValue) {
                reason = "hyperlink payload lengths outside BIFF8 limits";
                return false;
            }

            records.Add(new LegacyXlsHyperlinkRecord((ushort)BiffRecordType.HLink, hyperlinkPayload));
            if (!string.IsNullOrWhiteSpace(tooltip)) {
                if (tooltip!.Length > MaxTooltipCharacters) {
                    reason = "hyperlink tooltips outside BIFF8 limits";
                    return false;
                }

                byte[] tooltipPayload = BuildTooltipPayload(firstRow, firstColumn, lastRow, lastColumn, tooltip!);
                if (tooltipPayload.Length > ushort.MaxValue) {
                    reason = "hyperlink tooltips outside BIFF8 limits";
                    return false;
                }

                records.Add(new LegacyXlsHyperlinkRecord((ushort)BiffRecordType.HLinkTooltip, tooltipPayload));
            }

            return true;
        }

        private static byte[] BuildExternalUrlPayload(
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            string target,
            string? displayName,
            string? location) {
            using var stream = new MemoryStream();
            WriteRange(stream, firstRow, firstColumn, lastRow, lastColumn);
            WriteUInt32(stream, StreamVersion);
            uint flags = HasMoniker | IsAbsolute;
            if (!string.IsNullOrEmpty(displayName)) {
                flags |= HasDisplayName;
            }

            if (!string.IsNullOrWhiteSpace(location)) {
                flags |= HasLocation;
            }

            WriteUInt32(stream, flags);
            if (!string.IsNullOrEmpty(displayName)) {
                WriteHyperlinkString(stream, displayName!);
            }

            stream.Write(UrlMonikerClsid, 0, UrlMonikerClsid.Length);
            byte[] targetBytes = Encoding.Unicode.GetBytes(target + '\0');
            WriteUInt32(stream, checked((uint)targetBytes.Length));
            stream.Write(targetBytes, 0, targetBytes.Length);

            if (!string.IsNullOrWhiteSpace(location)) {
                WriteHyperlinkString(stream, NormalizeInternalLocation(location!));
            }

            return stream.ToArray();
        }

        private static byte[] BuildInternalLocationPayload(
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            string location,
            string? displayName) {
            using var stream = new MemoryStream();
            WriteRange(stream, firstRow, firstColumn, lastRow, lastColumn);
            WriteUInt32(stream, StreamVersion);
            uint flags = HasLocation;
            if (!string.IsNullOrEmpty(displayName)) {
                flags |= HasDisplayName;
            }

            WriteUInt32(stream, flags);
            if (!string.IsNullOrEmpty(displayName)) {
                WriteHyperlinkString(stream, displayName!);
            }

            WriteHyperlinkString(stream, location);
            return stream.ToArray();
        }

        private static byte[] BuildFileMonikerPayload(
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            string target,
            string? displayName,
            string? location) {
            using var stream = new MemoryStream();
            WriteRange(stream, firstRow, firstColumn, lastRow, lastColumn);
            WriteUInt32(stream, StreamVersion);
            uint flags = HasMoniker | IsAbsolute;
            if (!string.IsNullOrEmpty(displayName)) {
                flags |= HasDisplayName;
            }

            if (!string.IsNullOrWhiteSpace(location)) {
                flags |= HasLocation;
            }

            WriteUInt32(stream, flags);
            if (!string.IsNullOrEmpty(displayName)) {
                WriteHyperlinkString(stream, displayName!);
            }

            stream.Write(FileMonikerClsid, 0, FileMonikerClsid.Length);
            WriteUInt16(stream, 0);
            byte[] pathBytes = Encoding.ASCII.GetBytes(target + '\0');
            WriteUInt32(stream, checked((uint)pathBytes.Length));
            stream.Write(pathBytes, 0, pathBytes.Length);
            WriteUInt16(stream, 0xffff);
            WriteUInt16(stream, 0xdead);
            stream.Write(new byte[16], 0, 16);
            WriteUInt32(stream, 0);
            WriteUnicodeFileMonikerPath(stream, target);

            if (!string.IsNullOrWhiteSpace(location)) {
                WriteHyperlinkString(stream, NormalizeInternalLocation(location!));
            }

            return stream.ToArray();
        }

        private static byte[] BuildTooltipPayload(
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            string tooltip) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, (ushort)BiffRecordType.HLinkTooltip);
            WriteUInt16(stream, checked((ushort)(firstRow - 1)));
            WriteUInt16(stream, checked((ushort)(lastRow - 1)));
            WriteUInt16(stream, checked((ushort)(firstColumn - 1)));
            WriteUInt16(stream, checked((ushort)(lastColumn - 1)));
            byte[] tooltipBytes = Encoding.Unicode.GetBytes(tooltip + '\0');
            stream.Write(tooltipBytes, 0, tooltipBytes.Length);
            return stream.ToArray();
        }

        private static void WriteUnicodeFileMonikerPath(Stream stream, string target) {
            if (CanUseAnsiFileMonikerPath(target)) {
                WriteUInt32(stream, 0);
                return;
            }

            byte[] pathBytes = Encoding.Unicode.GetBytes(target + '\0');
            WriteUInt32(stream, checked((uint)(6 + pathBytes.Length)));
            WriteUInt32(stream, checked((uint)pathBytes.Length));
            WriteUInt16(stream, 3);
            stream.Write(pathBytes, 0, pathBytes.Length);
        }

        private static bool CanUseAnsiFileMonikerPath(string target) {
            for (int i = 0; i < target.Length; i++) {
                if (target[i] > 0x7f) {
                    return false;
                }
            }

            return true;
        }

        private static void WriteRange(Stream stream, int firstRow, int firstColumn, int lastRow, int lastColumn) {
            WriteUInt16(stream, checked((ushort)(firstRow - 1)));
            WriteUInt16(stream, checked((ushort)(lastRow - 1)));
            WriteUInt16(stream, checked((ushort)(firstColumn - 1)));
            WriteUInt16(stream, checked((ushort)(lastColumn - 1)));
            stream.Write(HLinkClsid, 0, HLinkClsid.Length);
        }

        private static bool TryParseBiff8Reference(string? reference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) {
            firstRow = firstColumn = lastRow = lastColumn = 0;
            string normalized = (reference ?? string.Empty).Replace("$", string.Empty);
            if (A1.TryParseRange(normalized, out firstRow, out firstColumn, out lastRow, out lastColumn)) {
                return IsInBiff8Grid(firstRow, firstColumn) && IsInBiff8Grid(lastRow, lastColumn);
            }

            (int row, int column) = A1.ParseCellRef(normalized);
            if (!IsInBiff8Grid(row, column)) {
                return false;
            }

            firstRow = lastRow = row;
            firstColumn = lastColumn = column;
            return true;
        }

        private static bool IsInBiff8Grid(int row, int column) {
            return row >= 1 && row <= Biff8MaxRow && column >= 1 && column <= Biff8MaxColumn;
        }

        private static string NormalizeInternalLocation(string location) {
            string normalized = location.Trim();
            return normalized.StartsWith("#", StringComparison.Ordinal)
                ? normalized.Substring(1)
                : normalized;
        }

        private static bool IsSupportedAbsoluteExternalUri(string target) {
            if (!Uri.TryCreate(target.Trim(), UriKind.Absolute, out Uri? uri)) {
                return false;
            }

            return string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                || string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)
                || string.Equals(uri.Scheme, Uri.UriSchemeFtp, StringComparison.OrdinalIgnoreCase)
                || string.Equals(uri.Scheme, Uri.UriSchemeMailto, StringComparison.OrdinalIgnoreCase)
                || string.Equals(uri.Scheme, Uri.UriSchemeFile, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsSupportedRelativeFileTarget(string target) {
            string trimmed = target.Trim();
            if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith("#", StringComparison.Ordinal)) {
                return false;
            }

            if (trimmed.StartsWith("/", StringComparison.Ordinal) || trimmed.StartsWith("\\", StringComparison.Ordinal)) {
                return false;
            }

            int colonIndex = trimmed.IndexOf(':');
            if (colonIndex >= 0) {
                int slashIndex = trimmed.IndexOfAny(new[] { '/', '\\' });
                if (slashIndex < 0 || colonIndex < slashIndex) {
                    return false;
                }
            }

            return Uri.TryCreate(trimmed.Replace('\\', '/'), UriKind.Relative, out _);
        }

        private static void WriteHyperlinkString(Stream stream, string value) {
            byte[] valueBytes = Encoding.Unicode.GetBytes(value + '\0');
            WriteUInt32(stream, checked((uint)(value.Length + 1)));
            stream.Write(valueBytes, 0, valueBytes.Length);
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
            stream.WriteByte((byte)((value >> 16) & 0xff));
            stream.WriteByte((byte)((value >> 24) & 0xff));
        }

        internal readonly struct LegacyXlsHyperlinkRecord {
            internal LegacyXlsHyperlinkRecord(ushort recordType, byte[] payload) {
                RecordType = recordType;
                Payload = payload;
            }

            internal ushort RecordType { get; }

            internal byte[] Payload { get; }
        }
    }
}
