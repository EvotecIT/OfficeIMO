using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Package;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Validates worksheet hyperlinks and builds their BIFF12 records and package relationships.</summary>
    internal sealed class XlsbWorksheetHyperlinkPlan {
        private const int BrtHLink = 494;
        private const string HyperlinkRelationshipSuffix = "/hyperlink";
        private const int MaximumLocationLengthExclusive = 2_084;
        private const int MaximumTooltipLengthExclusive = 256;

        private XlsbWorksheetHyperlinkPlan(
            IReadOnlyList<XlsbGeneratedRecord> records,
            IReadOnlyList<XlsbHyperlinkRelationship> relationships) {
            Records = records;
            Relationships = relationships;
        }

        internal IReadOnlyList<XlsbGeneratedRecord> Records { get; }

        internal IReadOnlyList<XlsbHyperlinkRelationship> Relationships { get; }

        internal bool RelationshipsMatch(XlsbWorksheet sourceSheet) {
            if (sourceSheet == null) throw new ArgumentNullException(nameof(sourceSheet));
            XlsbHyperlink[] sourceRelationships = sourceSheet.Hyperlinks
                .Where(hyperlink => hyperlink.IsExternal)
                .ToArray();
            if (sourceRelationships.Length != Relationships.Count) return false;
            var actual = Relationships.ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);
            foreach (XlsbHyperlink source in sourceRelationships) {
                if (!actual.TryGetValue(source.RelationshipId, out XlsbHyperlinkRelationship? relationship)
                    || !string.Equals(relationship.Target, source.ExternalTarget, StringComparison.Ordinal)) {
                    return false;
                }
            }
            return true;
        }

        internal byte[] CreateRelationshipPart(
            IReadOnlyDictionary<string, XlsbPackageRelationship>? sourceRelationships = null) {
            XlsbPackageRelationship[] preserved = sourceRelationships?.Values
                .Where(relationship => !IsHyperlinkRelationship(relationship))
                .ToArray() ?? Array.Empty<XlsbPackageRelationship>();
            var builder = new StringBuilder(160 + (Relationships.Count + preserved.Length) * 220);
            builder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            builder.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (XlsbPackageRelationship relationship in preserved) {
                AppendRelationship(
                    builder,
                    relationship.Id,
                    relationship.Type,
                    relationship.Target,
                    relationship.IsExternal);
            }
            foreach (XlsbHyperlinkRelationship relationship in Relationships) {
                AppendRelationship(
                    builder,
                    relationship.Id,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                    relationship.Target,
                    isExternal: true);
            }
            builder.Append("</Relationships>");
            return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(builder.ToString());
        }

        internal static XlsbWorksheetHyperlinkPlan Create(
            ExcelSheet sheet,
            bool pruneOrphanedRelationships = false,
            IReadOnlyCollection<string>? reservedRelationshipIds = null) {
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
                if (packageRelationships.Length != 0 && !pruneOrphanedRelationships) {
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
            var emittedRelationshipIds = new HashSet<string>(
                reservedRelationshipIds ?? Array.Empty<string>(),
                StringComparer.Ordinal);
            var emittedIdBySourceId = new Dictionary<string, string>(StringComparer.Ordinal);
            var records = new List<XlsbGeneratedRecord>(hyperlinks.Length);
            foreach (Hyperlink hyperlink in hyperlinks) {
                string sourceRelationshipId = hyperlink.Id?.Value ?? string.Empty;
                string emittedRelationshipId = sourceRelationshipId;
                if (sourceRelationshipId.Length != 0) {
                    if (!relationshipsById.ContainsKey(sourceRelationshipId)) {
                        throw new NotSupportedException($"Native XLSB generation found missing hyperlink relationship '{sourceRelationshipId}' on worksheet '{sheet.Name}'.");
                    }
                    usedRelationshipIds.Add(sourceRelationshipId);
                    if (!emittedIdBySourceId.TryGetValue(sourceRelationshipId, out emittedRelationshipId!)) {
                        emittedRelationshipId = emittedRelationshipIds.Add(sourceRelationshipId)
                            ? sourceRelationshipId
                            : AllocateRelationshipId(emittedRelationshipIds);
                        emittedIdBySourceId.Add(sourceRelationshipId, emittedRelationshipId);
                    }
                }
                records.Add(new XlsbGeneratedRecord(
                    BrtHLink,
                    CreatePayload(hyperlink, emittedRelationshipId, sheet.Name)));
            }

            if (usedRelationshipIds.Count != relationshipsById.Count && !pruneOrphanedRelationships) {
                throw new NotSupportedException($"Native XLSB generation found orphaned hyperlink relationships on worksheet '{sheet.Name}'.");
            }

            XlsbHyperlinkRelationship[] relationships = emittedIdBySourceId
                .OrderBy(pair => pair.Value, StringComparer.Ordinal)
                .Select(pair => new XlsbHyperlinkRelationship(
                    pair.Value,
                    relationshipsById[pair.Key].Uri.OriginalString))
                .ToArray();
            return new XlsbWorksheetHyperlinkPlan(records.AsReadOnly(), Array.AsReadOnly(relationships));
        }

        private static byte[] CreatePayload(
            Hyperlink hyperlink,
            string relationshipId,
            string sheetName) {
            EnsureOnlyAttributes(hyperlink, sheetName, "ref", "id", "location", "tooltip", "display");
            if (hyperlink.HasChildren) ThrowUnsupportedContent(hyperlink, sheetName);
            if (!TryParseRange(hyperlink.Reference?.Value, out XlsbCellRange? range)) {
                throw new NotSupportedException($"Native XLSB generation cannot encode hyperlink range '{hyperlink.Reference?.Value}' on worksheet '{sheetName}'.");
            }

            string location = hyperlink.Location?.Value ?? string.Empty;
            if (relationshipId.Length == 0 && string.IsNullOrWhiteSpace(location)) {
                throw new NotSupportedException($"Native XLSB generation found a hyperlink without an external target or internal location on worksheet '{sheetName}'.");
            }
            string tooltip = hyperlink.Tooltip?.Value ?? string.Empty;
            if (location.Length >= MaximumLocationLengthExclusive) {
                throw new NotSupportedException($"Native XLSB generation requires hyperlink locations shorter than 2,084 characters on worksheet '{sheetName}'.");
            }
            if (tooltip.Length >= MaximumTooltipLengthExclusive) {
                throw new NotSupportedException($"Native XLSB generation requires hyperlink tooltips shorter than 256 characters on worksheet '{sheetName}'.");
            }

            using var payload = new MemoryStream(64);
            WriteUInt32(payload, checked((uint)(range!.FirstRow - 1)));
            WriteUInt32(payload, checked((uint)(range.LastRow - 1)));
            WriteUInt32(payload, checked((uint)(range.FirstColumn - 1)));
            WriteUInt32(payload, checked((uint)(range.LastColumn - 1)));
            WriteWideString(payload, relationshipId);
            WriteWideString(payload, location);
            WriteWideString(payload, tooltip);
            WriteWideString(payload, hyperlink.Display?.Value ?? string.Empty);
            return payload.ToArray();
        }

        private static bool IsHyperlinkRelationship(XlsbPackageRelationship relationship) =>
            relationship.Type.EndsWith(HyperlinkRelationshipSuffix, StringComparison.Ordinal);

        private static string AllocateRelationshipId(ISet<string> usedIds) {
            for (int index = 1; ; index++) {
                string candidate = "rId" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (usedIds.Add(candidate)) return candidate;
            }
        }

        private static void AppendRelationship(
            StringBuilder builder,
            string id,
            string type,
            string target,
            bool isExternal) {
            builder.Append("<Relationship Id=\"");
            AppendXmlEscaped(builder, id);
            builder.Append("\" Type=\"");
            AppendXmlEscaped(builder, type);
            builder.Append("\" Target=\"");
            AppendXmlEscaped(builder, target);
            if (isExternal) builder.Append("\" TargetMode=\"External");
            builder.Append("\"/>");
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

        private static void AppendXmlEscaped(StringBuilder builder, string value) {
            foreach (char character in value) {
                switch (character) {
                    case '&': builder.Append("&amp;"); break;
                    case '<': builder.Append("&lt;"); break;
                    case '>': builder.Append("&gt;"); break;
                    case '"': builder.Append("&quot;"); break;
                    case '\'': builder.Append("&apos;"); break;
                    default: builder.Append(character); break;
                }
            }
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
