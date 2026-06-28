using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsDataConsolidationWriter {
        private const string ExternalLinkPathRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";

        internal static bool SupportsWorksheetDataConsolidation(ExcelSheet sheet, out string? reason) {
            reason = null;
            return TryCreatePayload(sheet, out _, out reason);
        }

        internal static bool TryCreatePayload(ExcelSheet sheet, out byte[]? payload) {
            return TryCreatePayload(sheet, out payload, out _);
        }

        internal static IReadOnlyList<byte[]> CreateReferencePayloads(ExcelSheet sheet) {
            DataConsolidate? dataConsolidate = sheet.WorksheetPart.Worksheet?.Elements<DataConsolidate>().FirstOrDefault();
            if (dataConsolidate == null) {
                return Array.Empty<byte[]>();
            }

            var payloads = new List<byte[]>();
            foreach (DataReference dataReference in dataConsolidate.Elements<DataReferences>().SelectMany(references => references.Elements<DataReference>())) {
                if (TryCreateReferencePayload(sheet, dataReference, out byte[]? payload, out _)) {
                    payloads.Add(payload!);
                }
            }

            return payloads;
        }

        internal static IReadOnlyList<byte[]> CreateNamePayloads(ExcelSheet sheet) {
            DataConsolidate? dataConsolidate = sheet.WorksheetPart.Worksheet?.Elements<DataConsolidate>().FirstOrDefault();
            if (dataConsolidate == null) {
                return Array.Empty<byte[]>();
            }

            var payloads = new List<byte[]>();
            foreach (DataReference dataReference in dataConsolidate.Elements<DataReferences>().SelectMany(references => references.Elements<DataReference>())) {
                if (TryCreateNamePayload(sheet, dataReference, out byte[]? payload, out _)) {
                    payloads.Add(payload!);
                }
            }

            return payloads;
        }

        private static bool TryCreatePayload(ExcelSheet sheet, out byte[]? payload, out string? reason) {
            payload = null;
            reason = null;
            DataConsolidate? dataConsolidate = sheet.WorksheetPart.Worksheet?.Elements<DataConsolidate>().FirstOrDefault();
            if (dataConsolidate == null) {
                return true;
            }

            if (!SupportsDataReferences(sheet, dataConsolidate, out reason)) {
                return false;
            }

            if (dataConsolidate.ChildElements.Any(child => child is not DataReferences)) {
                reason = "data consolidation source references";
                return false;
            }

            if (!SupportsKnownAttributes(dataConsolidate)) {
                reason = "data consolidation metadata";
                return false;
            }

            if (!TryGetFunction(dataConsolidate.Function?.Value ?? DataConsolidateFunctionValues.Sum, out ushort function)) {
                reason = "data consolidation functions";
                return false;
            }

            bool usesLeftLabels = dataConsolidate.LeftLabels?.Value == true || dataConsolidate.StartLabels?.Value == true;
            bool usesTopLabels = dataConsolidate.TopLabels?.Value == true;
            bool linksToSourceData = dataConsolidate.Link?.Value == true;

            payload = new byte[8];
            WriteUInt16(payload, 0, function);
            WriteUInt16(payload, 2, usesLeftLabels ? (ushort)1 : (ushort)0);
            WriteUInt16(payload, 4, usesTopLabels ? (ushort)1 : (ushort)0);
            WriteUInt16(payload, 6, linksToSourceData ? (ushort)1 : (ushort)0);
            return true;
        }

        private static bool SupportsDataReferences(ExcelSheet sheet, DataConsolidate dataConsolidate, out string? reason) {
            reason = null;
            foreach (DataReferences references in dataConsolidate.Elements<DataReferences>()) {
                if (!SupportsDataReferencesCollection(references)) {
                    reason = "data consolidation source references";
                    return false;
                }

                foreach (DataReference dataReference in references.Elements<DataReference>()) {
                    bool hasName = !string.IsNullOrWhiteSpace(dataReference.Name?.Value);
                    if (hasName) {
                        if (!TryCreateNamePayload(sheet, dataReference, out _, out reason)) {
                            return false;
                        }
                    } else if (!TryCreateReferencePayload(sheet, dataReference, out _, out reason)) {
                        return false;
                    }
                }
            }

            return true;
        }

        private static bool SupportsDataReferencesCollection(DataReferences references) {
            if (references.ChildElements.Any(child => child is not DataReference)) {
                return false;
            }

            foreach (OpenXmlAttribute attribute in references.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                if (!string.Equals(attribute.LocalName, "count", StringComparison.Ordinal)) {
                    return false;
                }
            }

            return true;
        }

        private static bool TryCreateReferencePayload(ExcelSheet sheet, DataReference dataReference, out byte[]? payload, out string? reason) {
            payload = null;
            reason = null;
            if (dataReference.HasChildren || !SupportsKnownReferenceAttributes(dataReference)) {
                reason = "data consolidation source references";
                return false;
            }

            if (!string.IsNullOrWhiteSpace(dataReference.Name?.Value)) {
                reason = "data consolidation named source references";
                return false;
            }

            string? relationshipId = GetRelationshipId(dataReference);
            bool hasExternalSource = !string.IsNullOrWhiteSpace(relationshipId);

            string? sheetName = dataReference.Sheet?.Value;
            string? reference = dataReference.Reference?.Value;
            if (!hasExternalSource && string.IsNullOrWhiteSpace(sheetName)) {
                reason = "data consolidation source references";
                return false;
            }

            if (hasExternalSource && !string.IsNullOrWhiteSpace(sheetName)) {
                reason = "external data consolidation sheet-qualified source references";
                return false;
            }

            if (string.IsNullOrWhiteSpace(reference)) {
                reason = "data consolidation source references";
                return false;
            }

            if (!TryParseRange(reference!, out ushort firstRow, out byte firstColumn, out ushort lastRow, out byte lastColumn)) {
                reason = "data consolidation source references outside BIFF8 worksheet limits";
                return false;
            }

            string source;
            if (hasExternalSource) {
                if (!TryGetExternalSource(sheet.WorksheetPart, relationshipId!, out string? externalSource, out reason)) {
                    return false;
                }

                source = ((char)0x01) + externalSource!;
            } else {
                source = ((char)0x02) + sheetName!;
            }

            if (source.Length > ushort.MaxValue) {
                reason = "data consolidation source references outside BIFF8 limits";
                return false;
            }

            if (GetReferencePayloadLength(source) > ushort.MaxValue) {
                reason = "data consolidation source reference payload lengths outside BIFF8 limits";
                return false;
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, firstRow);
            WriteUInt16(stream, lastRow);
            stream.WriteByte(firstColumn);
            stream.WriteByte(lastColumn);
            WriteUInt16(stream, checked((ushort)source.Length));
            WriteUnicodeStringNoCch(stream, source);
            payload = stream.ToArray();
            return true;
        }

        private static long GetReferencePayloadLength(string source) {
            return 8L + GetUnicodeStringNoCchPayloadLength(source);
        }

        private static long GetUnicodeStringNoCchPayloadLength(string text) {
            bool compressed = text.All(ch => ch <= byte.MaxValue);
            return 1L + (compressed ? text.Length : 2L * text.Length);
        }

        private static bool TryCreateNamePayload(ExcelSheet sheet, DataReference dataReference, out byte[]? payload, out string? reason) {
            payload = null;
            reason = null;
            if (dataReference.HasChildren || !SupportsKnownReferenceAttributes(dataReference)) {
                reason = "data consolidation source references";
                return false;
            }

            string? relationshipId = GetRelationshipId(dataReference);
            bool hasExternalSource = !string.IsNullOrWhiteSpace(relationshipId);

            string? name = dataReference.Name?.Value;
            if (string.IsNullOrWhiteSpace(name)) {
                reason = "data consolidation source references";
                return false;
            }

            if (name!.Length > byte.MaxValue) {
                reason = "data consolidation named references longer than 255 characters";
                return false;
            }

            if (!string.IsNullOrWhiteSpace(dataReference.Reference?.Value)) {
                reason = "named data consolidation source references with explicit cell ranges";
                return false;
            }

            string? sheetName = dataReference.Sheet?.Value;
            if (hasExternalSource && !string.IsNullOrWhiteSpace(sheetName)) {
                reason = "external data consolidation sheet-scoped named source references";
                return false;
            }

            string source;
            if (hasExternalSource) {
                if (!TryGetExternalSource(sheet.WorksheetPart, relationshipId!, out string? externalSource, out reason)) {
                    return false;
                }

                source = ((char)0x01) + externalSource!;
            } else {
                source = string.IsNullOrWhiteSpace(sheetName)
                    ? string.Empty
                    : ((char)0x02) + sheetName!;
            }

            if (source.Length > ushort.MaxValue) {
                reason = "data consolidation named references outside BIFF8 limits";
                return false;
            }

            if (GetNamePayloadLength(name, source) > ushort.MaxValue) {
                reason = "data consolidation named reference payload lengths outside BIFF8 limits";
                return false;
            }

            using var stream = new MemoryStream();
            WriteShortUnicodeString(stream, name);
            WriteUInt16(stream, checked((ushort)source.Length));
            if (source.Length > 0) {
                WriteUnicodeStringNoCch(stream, source);
            }

            payload = stream.ToArray();
            return true;
        }

        private static long GetNamePayloadLength(string name, string source) {
            long sourcePayloadLength = source.Length == 0 ? 0 : GetUnicodeStringNoCchPayloadLength(source);
            return 1L + GetUnicodeStringNoCchPayloadLength(name) + 2L + sourcePayloadLength;
        }

        private static bool TryParseRange(string reference, out ushort firstRow, out byte firstColumn, out ushort lastRow, out byte lastColumn) {
            firstRow = lastRow = 0;
            firstColumn = lastColumn = 0;
            string normalizedReference = reference.Trim().Replace("$", string.Empty);
            int row1;
            int column1;
            int row2;
            int column2;
            if (A1.TryParseRange(normalizedReference, out row1, out column1, out row2, out column2)) {
                return TryConvertRange(row1, column1, row2, column2, out firstRow, out firstColumn, out lastRow, out lastColumn);
            }

            if (A1.TryParseCellReferenceFast(normalizedReference, out row1, out column1)) {
                return TryConvertRange(row1, column1, row1, column1, out firstRow, out firstColumn, out lastRow, out lastColumn);
            }

            return false;
        }

        private static bool TryConvertRange(int row1, int column1, int row2, int column2, out ushort firstRow, out byte firstColumn, out ushort lastRow, out byte lastColumn) {
            firstRow = lastRow = 0;
            firstColumn = lastColumn = 0;
            if (row1 < 1 || row2 < row1 || row2 > 65536 || column1 < 1 || column2 < column1 || column2 > 256) {
                return false;
            }

            firstRow = checked((ushort)(row1 - 1));
            lastRow = checked((ushort)(row2 - 1));
            firstColumn = checked((byte)(column1 - 1));
            lastColumn = checked((byte)(column2 - 1));
            return true;
        }

        private static string? GetRelationshipId(OpenXmlElement element) {
            string? id = element.GetAttributes()
                .FirstOrDefault(attribute => string.Equals(attribute.LocalName, "id", StringComparison.Ordinal)
                    && !string.IsNullOrEmpty(attribute.NamespaceUri))
                .Value;
            return string.IsNullOrWhiteSpace(id)
                ? element.GetAttributes()
                    .FirstOrDefault(attribute => string.Equals(attribute.LocalName, "id", StringComparison.Ordinal)
                        && string.IsNullOrEmpty(attribute.NamespaceUri))
                    .Value
                : id;
        }

        private static bool TryGetExternalSource(WorksheetPart worksheetPart, string relationshipId, out string? source, out string? reason) {
            source = null;
            reason = null;
            ExternalRelationship? relationship = worksheetPart.ExternalRelationships
                .FirstOrDefault(item => string.Equals(item.Id, relationshipId, StringComparison.Ordinal));
            if (relationship == null) {
                reason = "external data consolidation source relationships";
                return false;
            }

            if (!string.Equals(relationship.RelationshipType, ExternalLinkPathRelationshipType, StringComparison.Ordinal)) {
                reason = "external data consolidation source relationship types";
                return false;
            }

            string target = relationship.Uri.OriginalString;
            if (string.IsNullOrWhiteSpace(target)) {
                reason = "external data consolidation source relationships";
                return false;
            }

            source = target;
            return true;
        }

        private static bool SupportsKnownAttributes(OpenXmlElement element) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                switch (attribute.LocalName) {
                    case "function":
                    case "leftLabels":
                    case "startLabels":
                    case "topLabels":
                    case "link":
                        break;
                    default:
                        return false;
                }
            }

            return true;
        }

        private static bool SupportsKnownReferenceAttributes(OpenXmlElement element) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (attribute.LocalName == "id") {
                    continue;
                }

                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                switch (attribute.LocalName) {
                    case "ref":
                    case "name":
                    case "sheet":
                    case "id":
                        break;
                    default:
                        return false;
                }
            }

            return true;
        }

        private static bool TryGetFunction(DataConsolidateFunctionValues function, out ushort value) {
            if (function == DataConsolidateFunctionValues.Average) {
                value = 0;
                return true;
            }

            if (function == DataConsolidateFunctionValues.CountNumbers) {
                value = 1;
                return true;
            }

            if (function == DataConsolidateFunctionValues.Count) {
                value = 2;
                return true;
            }

            if (function == DataConsolidateFunctionValues.Maximum) {
                value = 3;
                return true;
            }

            if (function == DataConsolidateFunctionValues.Minimum) {
                value = 4;
                return true;
            }

            if (function == DataConsolidateFunctionValues.Product) {
                value = 5;
                return true;
            }

            if (function == DataConsolidateFunctionValues.StandardDeviation) {
                value = 6;
                return true;
            }

            if (function == DataConsolidateFunctionValues.StandardDeviationP) {
                value = 7;
                return true;
            }

            if (function == DataConsolidateFunctionValues.Sum) {
                value = 8;
                return true;
            }

            if (function == DataConsolidateFunctionValues.Variance) {
                value = 9;
                return true;
            }

            if (function == DataConsolidateFunctionValues.VarianceP) {
                value = 10;
                return true;
            }

            value = 0;
            return false;
        }

        private static void WriteUInt16(byte[] buffer, int offset, ushort value) {
            buffer[offset] = (byte)(value & 0xff);
            buffer[offset + 1] = (byte)((value >> 8) & 0xff);
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private static void WriteShortUnicodeString(Stream stream, string text) {
            stream.WriteByte(checked((byte)text.Length));
            WriteUnicodeStringNoCch(stream, text);
        }

        private static void WriteUnicodeStringNoCch(Stream stream, string text) {
            bool compressed = text.All(ch => ch <= byte.MaxValue);
            stream.WriteByte(compressed ? (byte)0x00 : (byte)0x01);
            foreach (char ch in text) {
                if (compressed) {
                    stream.WriteByte((byte)(ch & 0xff));
                } else {
                    WriteUInt16(stream, ch);
                }
            }
        }
    }
}
