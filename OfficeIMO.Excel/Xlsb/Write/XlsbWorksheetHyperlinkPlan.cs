using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Validates worksheet hyperlinks and builds their BIFF12 records and package relationships.</summary>
    internal sealed class XlsbWorksheetHyperlinkPlan {
        private const int BrtHLink = 494;

        private XlsbWorksheetHyperlinkPlan(
            IReadOnlyList<XlsbGeneratedRecord> records,
            IReadOnlyList<XlsbHyperlinkRelationship> relationships) {
            Records = records;
            Relationships = relationships;
        }

        internal IReadOnlyList<XlsbGeneratedRecord> Records { get; }

        internal IReadOnlyList<XlsbHyperlinkRelationship> Relationships { get; }

        internal static XlsbWorksheetHyperlinkPlan Create(ExcelSheet sheet) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            Hyperlinks[] containers = worksheet.Elements<Hyperlinks>().ToArray();
            if (containers.Length > 1) {
                throw new NotSupportedException($"Native XLSB generation requires at most one hyperlinks element on worksheet '{sheet.Name}'.");
            }

            HyperlinkRelationship[] packageRelationships = sheet.WorksheetPart.HyperlinkRelationships.ToArray();
            var relationshipsById = new Dictionary<string, HyperlinkRelationship>(StringComparer.Ordinal);
            foreach (HyperlinkRelationship relationship in packageRelationships) {
                if (string.IsNullOrWhiteSpace(relationship.Id)
                    || !relationship.IsExternal
                    || relationshipsById.ContainsKey(relationship.Id)) {
                    throw new NotSupportedException($"Native XLSB generation found an invalid hyperlink relationship on worksheet '{sheet.Name}'.");
                }
                relationshipsById.Add(relationship.Id, relationship);
            }

            if (containers.Length == 0) {
                if (packageRelationships.Length != 0) {
                    throw new NotSupportedException($"Native XLSB generation found orphaned hyperlink relationships on worksheet '{sheet.Name}'.");
                }
                return new XlsbWorksheetHyperlinkPlan(
                    Array.Empty<XlsbGeneratedRecord>(),
                    Array.Empty<XlsbHyperlinkRelationship>());
            }

            Hyperlinks container = containers[0];
            EnsureOnlyAttributes(container, sheet.Name);
            Hyperlink[] hyperlinks = container.Elements<Hyperlink>().ToArray();
            if (hyperlinks.Length != container.ChildElements.Count) ThrowUnsupportedContent(container, sheet.Name);

            var usedRelationshipIds = new HashSet<string>(StringComparer.Ordinal);
            var records = new List<XlsbGeneratedRecord>(hyperlinks.Length);
            foreach (Hyperlink hyperlink in hyperlinks) {
                records.Add(new XlsbGeneratedRecord(
                    BrtHLink,
                    CreatePayload(hyperlink, relationshipsById, usedRelationshipIds, sheet.Name)));
            }

            if (usedRelationshipIds.Count != relationshipsById.Count) {
                throw new NotSupportedException($"Native XLSB generation found orphaned hyperlink relationships on worksheet '{sheet.Name}'.");
            }

            XlsbHyperlinkRelationship[] relationships = relationshipsById
                .Where(pair => usedRelationshipIds.Contains(pair.Key))
                .OrderBy(pair => pair.Key, StringComparer.Ordinal)
                .Select(pair => new XlsbHyperlinkRelationship(pair.Key, pair.Value.Uri.OriginalString))
                .ToArray();
            return new XlsbWorksheetHyperlinkPlan(records.AsReadOnly(), Array.AsReadOnly(relationships));
        }

        private static byte[] CreatePayload(
            Hyperlink hyperlink,
            IReadOnlyDictionary<string, HyperlinkRelationship> relationshipsById,
            ISet<string> usedRelationshipIds,
            string sheetName) {
            EnsureOnlyAttributes(hyperlink, sheetName, "ref", "id", "location", "tooltip", "display");
            if (hyperlink.HasChildren) ThrowUnsupportedContent(hyperlink, sheetName);
            if (!TryParseRange(hyperlink.Reference?.Value, out XlsbCellRange? range)) {
                throw new NotSupportedException($"Native XLSB generation cannot encode hyperlink range '{hyperlink.Reference?.Value}' on worksheet '{sheetName}'.");
            }

            string relationshipId = hyperlink.Id?.Value ?? string.Empty;
            string location = hyperlink.Location?.Value ?? string.Empty;
            if (relationshipId.Length != 0) {
                if (!relationshipsById.ContainsKey(relationshipId)) {
                    throw new NotSupportedException($"Native XLSB generation found missing hyperlink relationship '{relationshipId}' on worksheet '{sheetName}'.");
                }
                usedRelationshipIds.Add(relationshipId);
            } else if (string.IsNullOrWhiteSpace(location)) {
                throw new NotSupportedException($"Native XLSB generation found a hyperlink without an external target or internal location on worksheet '{sheetName}'.");
            }

            using var payload = new MemoryStream(64);
            WriteUInt32(payload, checked((uint)(range!.FirstRow - 1)));
            WriteUInt32(payload, checked((uint)(range.LastRow - 1)));
            WriteUInt32(payload, checked((uint)(range.FirstColumn - 1)));
            WriteUInt32(payload, checked((uint)(range.LastColumn - 1)));
            WriteWideString(payload, relationshipId);
            WriteWideString(payload, location);
            WriteWideString(payload, hyperlink.Tooltip?.Value ?? string.Empty);
            WriteWideString(payload, hyperlink.Display?.Value ?? string.Empty);
            return payload.ToArray();
        }

        private static bool TryParseRange(string? reference, out XlsbCellRange? range) {
            range = null;
            if (string.IsNullOrWhiteSpace(reference)) return false;
            if (A1.TryParseRange(reference!, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                range = new XlsbCellRange(firstRow, lastRow, firstColumn, lastColumn);
                return true;
            }
            if (!A1.TryParseCellReferenceFast(reference!, out firstRow, out firstColumn)) return false;
            range = new XlsbCellRange(firstRow, firstRow, firstColumn, firstColumn);
            return true;
        }

        private static void EnsureOnlyAttributes(OpenXmlElement element, string sheetName, params string[] allowedNames) {
            var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
            OpenXmlAttribute? unsupported = element.GetAttributes()
                .Cast<OpenXmlAttribute?>()
                .FirstOrDefault(attribute => attribute.HasValue
                    && !string.Equals(attribute.Value.NamespaceUri, "http://www.w3.org/2000/xmlns/", StringComparison.Ordinal)
                    && !allowed.Contains(attribute.Value.LocalName));
            if (unsupported.HasValue) {
                throw new NotSupportedException($"Native XLSB generation does not yet support attribute '{unsupported.Value.LocalName}' on worksheet element '{element.LocalName}' in worksheet '{sheetName}'.");
            }
        }

        private static void ThrowUnsupportedContent(OpenXmlElement element, string sheetName) =>
            throw new NotSupportedException($"Native XLSB generation does not yet support child content in worksheet element '{element.LocalName}' on worksheet '{sheetName}'.");

        private static void WriteWideString(Stream output, string value) {
            WriteUInt32(output, checked((uint)value.Length));
            byte[] bytes = Encoding.Unicode.GetBytes(value);
            output.Write(bytes, 0, bytes.Length);
        }

        private static void WriteUInt32(Stream output, uint value) {
            output.WriteByte((byte)value);
            output.WriteByte((byte)(value >> 8));
            output.WriteByte((byte)(value >> 16));
            output.WriteByte((byte)(value >> 24));
        }
    }

    /// <summary>Describes one external worksheet hyperlink relationship.</summary>
    internal sealed class XlsbHyperlinkRelationship {
        internal XlsbHyperlinkRelationship(string id, string target) {
            Id = id ?? throw new ArgumentNullException(nameof(id));
            Target = target ?? throw new ArgumentNullException(nameof(target));
        }

        internal string Id { get; }

        internal string Target { get; }
    }
}
