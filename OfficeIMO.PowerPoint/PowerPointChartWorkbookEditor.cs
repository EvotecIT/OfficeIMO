using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Decides whether a chart's embedded workbook is simple enough to rewrite, and rewrites its single data sheet, on
    /// OfficeIMO's own package layer.
    /// </summary>
    internal static class PowerPointChartWorkbookEditor {
        private const string RelationshipTypePrefix = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
        private const string StrictRelationshipTypePrefix = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
        private const string OfficeDocumentRelationshipType = RelationshipTypePrefix + "officeDocument";
        private const string CorePropertiesRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        private const string ExtendedPropertiesRelationshipType = RelationshipTypePrefix + "extended-properties";
        private const string CustomPropertiesRelationshipType = RelationshipTypePrefix + "custom-properties";
        private const string WorksheetRelationshipType = RelationshipTypePrefix + "worksheet";
        private const string SharedStringsRelationshipType = RelationshipTypePrefix + "sharedStrings";
        private const string StylesRelationshipType = RelationshipTypePrefix + "styles";
        private const string ThemeRelationshipType = RelationshipTypePrefix + "theme";
        private const string TableRelationshipType = RelationshipTypePrefix + "table";
        private const string HyperlinkRelationshipType = RelationshipTypePrefix + "hyperlink";
        private const string MediaRelationshipType = "http://schemas.microsoft.com/office/2007/relationships/media";
        private const string AudioRelationshipType = RelationshipTypePrefix + "audio";
        private const string VideoRelationshipType = RelationshipTypePrefix + "video";

        // Main-part content types of the other spreadsheet document kinds. Any other content type opens as a plain
        // workbook, as the SDK's document type reads it.
        private static readonly HashSet<string> OtherSpreadsheetContentTypes = new(StringComparer.Ordinal) {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            "application/vnd.ms-excel.template.macroEnabled.main+xml",
            "application/vnd.ms-excel.addin.macroEnabled.main+xml"
        };
        private const string WorksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        private const string SharedStringsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        private const string StylesContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        private const string ThemeContentType = "application/vnd.openxmlformats-officedocument.theme+xml";
        private const string TableContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
        private const string CorePropertiesContentType = "application/vnd.openxmlformats-package.core-properties+xml";
        private const string ExtendedPropertiesContentType = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        private const string CustomPropertiesContentType = "application/vnd.openxmlformats-officedocument.custom-properties+xml";

        private const int SpreadsheetML = OfficeXmlNamespaces.SpreadsheetML;

        // Worksheet content whose meaning depends on cell positions, formulas or other parts.
        private static readonly HashSet<string> MeaningChangingWorksheetElements = new(StringComparer.Ordinal) {
            "f", "hyperlinks", "oleObjects", "controls", "mergeCells", "conditionalFormatting", "dataValidations",
            "sheetProtection", "protectedRanges", "scenarios", "customSheetViews", "pivotTableDefinition", "drawing",
            "legacyDrawing"
        };

        /// <summary>
        /// Whether the workbook is a plain workbook with one sheet named Sheet1, shared strings, optional styles, theme
        /// and one simple table, and nothing whose meaning a rewrite of the cells could change.
        /// </summary>
        /// <exception cref="InvalidDataException">The bytes are not a readable package.</exception>
        internal static bool IsSafelyEditable(byte[] workbook) {
            ArgumentNullException.ThrowIfNull(workbook);
            using var stream = new MemoryStream(workbook, writable: false);
            using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(stream, leaveOpen: true);
            return TryGetEditableParts(package, out _);
        }

        internal static byte[] Update(byte[] source,
            OfficeChartData data) {
            ArgumentNullException.ThrowIfNull(source);
            ArgumentNullException.ThrowIfNull(data);
            using var input = new MemoryStream(source, writable: false);
            using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(input, leaveOpen: true);
            if (!TryGetEditableParts(package, out EditableParts parts)) {
                throw new NotSupportedException(
                    "The imported workbook contains sheets, formulas, macros, relationships, or range-dependent content that cannot be updated without changing its meaning.");
            }
            SpreadsheetMarkup markup = parts.Markup;

            OfficeXmlTree worksheetTree = package.GetXml(parts.Worksheet);
            var worksheet = new OfficeXmlNode(worksheetTree, worksheetTree.Root);
            OfficeXmlNode sheetData = markup.Child(worksheet, "sheetData");
            if (sheetData.IsNone) sheetData = worksheet.Append(worksheet.CreateElement(SpreadsheetML, "sheetData"));
            OfficeXmlNode table = default;
            if (parts.Table != null) {
                OfficeXmlTree tableTree = package.GetXml(parts.Table);
                table = new OfficeXmlNode(tableTree, tableTree.Root);
                int columnCount = markup.Children(markup.Child(table, "tableColumns"), "tableColumn").Count;
                if (columnCount != data.Series.Count + 1) {
                    throw new NotSupportedException(
                        "Changing the series count of a producer-authored chart table can alter its meaning.");
                }
            }

            OfficeXmlNode sharedStrings = markup.Root(package, parts.SharedStrings, "sst");
            if (sharedStrings.IsNone)
                throw new InvalidDataException("The shared strings part '" + parts.SharedStrings + "' has no shared string table.");
            var indices = new Dictionary<string, int>(
                StringComparer.Ordinal);
            int sharedStringIndex = 0;
            foreach (OfficeXmlNode item in markup.Children(sharedStrings, "si")) {
                string value = GetSharedStringText(markup, item);
                if (!indices.ContainsKey(value))
                    indices.Add(value, sharedStringIndex);
                sharedStringIndex++;
            }

            int GetSharedStringIndex(string? value) {
                string text = value ?? string.Empty;
                if (indices.TryGetValue(text, out int existing))
                    return existing;
                int created = sharedStringIndex++;
                OfficeXmlNode item = sharedStrings.Append(sharedStrings.CreateElement(SpreadsheetML, "si"));
                OfficeXmlNode textElement = item.Append(item.CreateElement(SpreadsheetML, "t"));
                if (text.Length != text.Trim().Length)
                    textElement.SetAttribute(OfficeXmlNamespaces.Xml, "space", "preserve");
                if (text.Length > 0) textElement.SetInnerText(text);
                indices.Add(text, created);
                return created;
            }

            List<OfficeXmlNode> existingRows = markup.Children(sheetData, "row");
            Dictionary<string, OfficeXmlNode> existingCells = existingRows
                .SelectMany(row => markup.Children(row, "c"))
                .Where(cell => !string.IsNullOrWhiteSpace(cell.Attribute("r")))
                .ToDictionary(cell => cell.Attribute("r")!,
                    cell => cell, StringComparer.OrdinalIgnoreCase);

            OfficeXmlNode CreateRow(uint rowIndex, bool headerRow) {
                OfficeXmlNode template = existingRows.FirstOrDefault(row =>
                    GetRowIndex(row) == rowIndex);
                if (template.IsNone) {
                    template = headerRow
                        ? existingRows.FirstOrDefault(row => GetRowIndex(row) == 1U)
                        : existingRows.FirstOrDefault(row => GetRowIndex(row) >= 2U);
                }
                OfficeXmlNode row = template.IsNone
                    ? sheetData.CreateElement(SpreadsheetML, "row")
                    : CopyWithoutChildren(template);
                row.SetAttribute("r", rowIndex.ToString(CultureInfo.InvariantCulture));
                return row;
            }

            OfficeXmlNode CreateCell(string reference, string column,
                bool headerCell) {
                if (!existingCells.TryGetValue(reference,
                        out OfficeXmlNode template)) {
                    string preferred = column + (headerCell ? "1" : "2");
                    existingCells.TryGetValue(preferred, out template);
                }
                // A copy carries the template's attributes only, so it has no value or inline string yet.
                OfficeXmlNode cell = template.IsNone
                    ? sheetData.CreateElement(SpreadsheetML, "c")
                    : CopyWithoutChildren(template);
                cell.SetAttribute("r", reference);
                return cell;
            }

            foreach (OfficeXmlNode row in existingRows) row.Remove();
            OfficeXmlNode header = CreateRow(1U, headerRow: true);
            AppendSharedStringCell(header,
                CreateCell("A1", "A", headerCell: true),
                GetSharedStringIndex(" "));
            for (int seriesIndex = 0;
                 seriesIndex < data.Series.Count; seriesIndex++) {
                string column = GetExcelColumn(seriesIndex + 2);
                AppendSharedStringCell(header,
                    CreateCell(column + "1", column,
                        headerCell: true),
                    GetSharedStringIndex(data.Series[seriesIndex].Name));
            }
            sheetData.Append(header);

            for (int rowIndex = 0;
                 rowIndex < data.Categories.Count; rowIndex++) {
                uint excelRow = checked((uint)rowIndex + 2U);
                string excelRowText = excelRow.ToString(CultureInfo.InvariantCulture);
                OfficeXmlNode row = CreateRow(excelRow, headerRow: false);
                AppendSharedStringCell(row,
                    CreateCell("A" + excelRowText, "A",
                        headerCell: false),
                    GetSharedStringIndex(data.Categories[rowIndex]));
                for (int seriesIndex = 0;
                     seriesIndex < data.Series.Count; seriesIndex++) {
                    string column = GetExcelColumn(seriesIndex + 2);
                    OfficeXmlNode cell = row.Append(CreateCell(column + excelRowText, column,
                        headerCell: false));
                    cell.RemoveAttribute("t");
                    AppendValue(cell, FormatNumber(data.Series[seriesIndex].Values[rowIndex]));
                }
                sheetData.Append(row);
            }

            string lastColumn = GetExcelColumn(data.Series.Count + 1);
            uint lastRow = checked((uint)data.Categories.Count + 1U);
            string reference = "A1:" + lastColumn + lastRow.ToString(CultureInfo.InvariantCulture);
            OfficeXmlNode dimension = markup.Child(worksheet, "dimension");
            if (dimension.IsNone)
                dimension = worksheet.InsertBefore(worksheet.CreateElement(SpreadsheetML, "dimension"), worksheet.FirstChild);
            dimension.SetAttribute("ref", reference);
            if (!table.IsNone) {
                table.SetAttribute("ref", reference);
                OfficeXmlNode autoFilter = markup.Child(table, "autoFilter");
                if (!autoFilter.IsNone)
                    autoFilter.SetAttribute("ref", reference);
                List<OfficeXmlNode> columns = markup.Children(markup.Child(table, "tableColumns"), "tableColumn");
                columns[0].SetAttribute("name", " ");
                for (int index = 0; index < data.Series.Count; index++)
                    columns[index + 1].SetAttribute("name", data.Series[index].Name);
            }
            sharedStrings.SetAttribute("count", checked((uint)(1
                + data.Series.Count + data.Categories.Count)).ToString(CultureInfo.InvariantCulture));
            sharedStrings.SetAttribute("uniqueCount", markup.Children(sharedStrings, "si").Count
                .ToString(CultureInfo.InvariantCulture));

            using var output = new MemoryStream();
            package.Save(output);
            return output.ToArray();
        }

        private readonly struct EditableParts {
            internal EditableParts(SpreadsheetMarkup markup, string worksheet, string sharedStrings, string? table) {
                Markup = markup;
                Worksheet = worksheet;
                SharedStrings = sharedStrings;
                Table = table;
            }

            internal SpreadsheetMarkup Markup { get; }
            internal string Worksheet { get; }
            internal string SharedStrings { get; }
            internal string? Table { get; }
        }

        /// <summary>
        /// Finds SpreadsheetML elements as the SDK reads them: Strict-namespace elements count only in a package whose
        /// main relationship is Strict; elsewhere they are foreign markup.
        /// </summary>
        private readonly struct SpreadsheetMarkup {
            private readonly bool _strict;

            internal SpreadsheetMarkup(bool strict) => _strict = strict;

            internal bool Is(OfficeXmlNode element, string localName) =>
                element.Is(SpreadsheetML, localName) && IsReadable(element);

            internal bool IsAny(OfficeXmlNode element, HashSet<string> localNames) =>
                element.NamespaceId == SpreadsheetML && localNames.Contains(element.LocalName) && IsReadable(element);

            private bool IsReadable(OfficeXmlNode element) =>
                _strict || element.Tree == null || !OfficeXmlNamespaces.IsStrictUriId(element.Tree.UriIdOf(element.Node));

            // The part's root element when it has the expected name; none otherwise.
            internal OfficeXmlNode Root(OfficeOpcEditablePackage package, string partName, string localName) {
                OfficeXmlTree tree = package.GetXml(partName);
                var root = new OfficeXmlNode(tree, tree.Root);
                return Is(root, localName) ? root : default;
            }

            internal OfficeXmlNode Child(OfficeXmlNode parent, string localName) {
                foreach (OfficeXmlNode child in parent.Children(SpreadsheetML, localName)) {
                    if (IsReadable(child)) return child;
                }
                return default;
            }

            internal List<OfficeXmlNode> Children(OfficeXmlNode parent, string localName) {
                var children = new List<OfficeXmlNode>();
                foreach (OfficeXmlNode child in parent.Children(SpreadsheetML, localName)) {
                    if (IsReadable(child)) children.Add(child);
                }
                return children;
            }

            internal List<OfficeXmlNode> Descendants(OfficeXmlNode element, string localName) =>
                element.IsNone ? new List<OfficeXmlNode>() : element.Descendants(SpreadsheetML, localName).Where(IsReadable).ToList();

            internal bool ContainsAny(OfficeXmlNode element, HashSet<string> localNames) {
                foreach (OfficeXmlNode child in element.Children()) {
                    if (IsAny(child, localNames) || ContainsAny(child, localNames)) return true;
                }
                return false;
            }
        }

        private enum RelationshipKind {
            Part,
            External,
            Hyperlink,
            DataPartReference
        }

        private readonly struct Relationship {
            internal Relationship(string id, string type, bool strict, RelationshipKind kind, string? targetPart) {
                Id = id;
                Type = type;
                IsStrict = strict;
                Kind = kind;
                TargetPart = targetPart;
            }

            internal string Id { get; }

            /// <summary>The relationship type, Strict types given their transitional spelling.</summary>
            internal string Type { get; }

            internal bool IsStrict { get; }
            internal RelationshipKind Kind { get; }

            /// <summary>The target part of an internal relationship, or null when the target names no part.</summary>
            internal string? TargetPart { get; }
        }

        private static bool TryGetEditableParts(OfficeOpcEditablePackage package, out EditableParts parts) {
            parts = default;

            // The package relates to the workbook once and otherwise only to document properties.
            string? workbookPart = null;
            bool strict = false;
            int workbookCount = 0;
            foreach (Relationship relationship in ReadRelationships(package, null)) {
                if (relationship.Kind == RelationshipKind.External) return false;
                if (relationship.Kind != RelationshipKind.Part) continue;
                switch (relationship.Type) {
                    case OfficeDocumentRelationshipType:
                        workbookCount++;
                        workbookPart = relationship.TargetPart;
                        strict = relationship.IsStrict;
                        break;
                    case CorePropertiesRelationshipType:
                        if (!HasContentType(package, relationship.TargetPart, CorePropertiesContentType)) return false;
                        break;
                    case ExtendedPropertiesRelationshipType:
                        if (!HasContentType(package, relationship.TargetPart, ExtendedPropertiesContentType)) return false;
                        break;
                    case CustomPropertiesRelationshipType:
                        if (!HasContentType(package, relationship.TargetPart, CustomPropertiesContentType)) return false;
                        break;
                    default:
                        return false;
                }
            }
            if (workbookCount != 1 || workbookPart == null || !package.PartExists(workbookPart)) return false;
            string? workbookContentType = package.GetContentType(workbookPart);
            if (workbookContentType == null || OtherSpreadsheetContentTypes.Contains(workbookContentType)) return false;
            var markup = new SpreadsheetMarkup(strict);
            OfficeXmlNode workbook = markup.Root(package, workbookPart, "workbook");
            if (workbook.IsNone) return false;

            // The workbook relates only to one worksheet, the shared strings and at most one style sheet and theme.
            var worksheets = new List<Relationship>();
            var sharedStrings = new List<Relationship>();
            var styles = new List<Relationship>();
            var themes = new List<Relationship>();
            foreach (Relationship relationship in ReadRelationships(package, workbookPart)) {
                if (relationship.Kind != RelationshipKind.Part) return false;
                switch (relationship.Type) {
                    case WorksheetRelationshipType: worksheets.Add(relationship); break;
                    case SharedStringsRelationshipType: sharedStrings.Add(relationship); break;
                    case StylesRelationshipType: styles.Add(relationship); break;
                    case ThemeRelationshipType: themes.Add(relationship); break;
                    default: return false;
                }
            }
            if (worksheets.Count != 1 || sharedStrings.Count != 1 || styles.Count > 1 || themes.Count > 1
                || !HasContentType(package, worksheets[0].TargetPart, WorksheetContentType)
                || !HasContentType(package, sharedStrings[0].TargetPart, SharedStringsContentType)
                || styles.Any(style => !HasContentType(package, style.TargetPart, StylesContentType))
                || themes.Any(theme => !HasContentType(package, theme.TargetPart, ThemeContentType))) {
                return false;
            }

            List<OfficeXmlNode> sheets = markup.Children(markup.Child(workbook, "sheets"), "sheet");
            if (sheets.Count != 1) return false;
            OfficeXmlNode sheet = sheets[0];
            string? sheetId = sheet.Attribute(OfficeXmlNamespaces.Relationships, "id");
            if (sheetId == null
                || !string.Equals(sheetId, worksheets[0].Id, StringComparison.Ordinal)
                || !string.Equals(sheet.Attribute("name"), "Sheet1", StringComparison.Ordinal)
                || !markup.Child(workbook, "definedNames").IsNone
                || !markup.Child(workbook, "externalReferences").IsNone) {
                return false;
            }

            // The worksheet relates to at most one table, and the parts below the workbook relate to nothing.
            string worksheetPart = worksheets[0].TargetPart!;
            OfficeXmlNode worksheet = markup.Root(package, worksheetPart, "worksheet");
            if (worksheet.IsNone) return false;
            var tables = new List<Relationship>();
            foreach (Relationship relationship in ReadRelationships(package, worksheetPart)) {
                if (relationship.Kind != RelationshipKind.Part || relationship.Type != TableRelationshipType) return false;
                tables.Add(relationship);
            }
            if (tables.Count > 1
                || tables.Any(table => !HasContentType(package, table.TargetPart, TableContentType))
                || HasRelationships(package, sharedStrings[0].TargetPart!)
                || styles.Any(style => HasRelationships(package, style.TargetPart!))
                || themes.Any(theme => HasRelationships(package, theme.TargetPart!))
                || tables.Any(table => HasRelationships(package, table.TargetPart!))
                || markup.ContainsAny(worksheet, MeaningChangingWorksheetElements)) {
                return false;
            }

            string? tablePart = tables.Count == 0 ? null : tables[0].TargetPart;
            if (tablePart != null && !IsSimpleChartDataTable(markup, markup.Root(package, tablePart, "table"))) return false;
            parts = new EditableParts(markup, worksheetPart, sharedStrings[0].TargetPart!, tablePart);
            return true;
        }

        // The relationships of a part (null for the package), read without creating a missing relationships part.
        private static List<Relationship> ReadRelationships(OfficeOpcEditablePackage package, string? sourcePartName) {
            var relationships = new List<Relationship>();
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(sourcePartName);
            if (!package.PartExists(relationshipsPart)) return relationships;
            OfficeXmlTree tree = package.GetXml(relationshipsPart);
            var root = new OfficeXmlNode(tree, tree.Root);
            foreach (OfficeXmlNode element in root.Children(OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                string id = element.Attribute("Id") ?? string.Empty;
                string rawType = element.Attribute("Type") ?? string.Empty;
                string type = NormalizeRelationshipType(rawType);
                string? target = element.Attribute("Target");
                bool external = string.Equals(element.Attribute("TargetMode")?.Trim(), "External", StringComparison.Ordinal);
                RelationshipKind kind;
                if (external) {
                    kind = type == HyperlinkRelationshipType ? RelationshipKind.Hyperlink : RelationshipKind.External;
                } else if (type == HyperlinkRelationshipType) {
                    kind = RelationshipKind.Hyperlink;
                } else if (type is MediaRelationshipType or AudioRelationshipType or VideoRelationshipType) {
                    kind = RelationshipKind.DataPartReference;
                } else {
                    kind = RelationshipKind.Part;
                }
                string? targetPart = kind == RelationshipKind.Part
                    ? OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, target)
                    : null;
                relationships.Add(new Relationship(id, type, rawType.StartsWith(StrictRelationshipTypePrefix, StringComparison.Ordinal), kind, targetPart));
            }
            return relationships;
        }

        // Strict relationship types in their transitional spelling. The two document-property types are spelled
        // differently in Strict; their transitional spelling under the Strict prefix is no known type.
        private static string NormalizeRelationshipType(string type) {
            if (!type.StartsWith(StrictRelationshipTypePrefix, StringComparison.Ordinal)) return type;
            switch (type.Substring(StrictRelationshipTypePrefix.Length)) {
                case "extendedProperties": return ExtendedPropertiesRelationshipType;
                case "customProperties": return CustomPropertiesRelationshipType;
                case "extended-properties":
                case "custom-properties": return type;
                default: return OfficeXmlNamespaces.NormalizeRelationshipType(type);
            }
        }

        private static bool HasRelationships(OfficeOpcEditablePackage package, string partName) {
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(partName);
            if (!package.PartExists(relationshipsPart)) return false;
            OfficeXmlTree tree = package.GetXml(relationshipsPart);
            return tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship") != OfficeXmlTree.NoNode;
        }

        // Parts of a fixed type must carry exactly its content type.
        private static bool HasContentType(OfficeOpcEditablePackage package, string? partName, string contentType) =>
            partName != null && package.PartExists(partName)
            && string.Equals(package.GetContentType(partName), contentType, StringComparison.Ordinal);

        // The text of a shared string item as a reader of the item sees it: the text of its plain, rich and phonetic
        // runs in document order.
        private static string GetSharedStringText(SpreadsheetMarkup markup, OfficeXmlNode item) {
            List<OfficeXmlNode> texts = markup.Descendants(item, "t");
            if (texts.Count == 1) return texts[0].InnerText;
            var builder = new StringBuilder();
            foreach (OfficeXmlNode text in texts) builder.Append(text.InnerText);
            return builder.ToString();
        }

        private static uint? GetRowIndex(OfficeXmlNode row) =>
            uint.TryParse(row.Attribute("r"), NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite,
                CultureInfo.InvariantCulture, out uint value) ? value : null;

        // A copy of the element with its attributes and namespace declarations but no children.
        private static OfficeXmlNode CopyWithoutChildren(OfficeXmlNode element) {
            OfficeXmlTree tree = element.Tree!;
            int copy = tree.CreateElement(element.NamespaceId, element.LocalName);
            int count = tree.AttributeCountOf(element.Node);
            for (int index = 0; index < count; index++) {
                tree.GetAttributeAt(element.Node, index, out int namespaceId, out string localName, out string value);
                tree.SetAttribute(copy, namespaceId, localName, value);
            }
            return new OfficeXmlNode(tree, copy);
        }

        private static void AppendSharedStringCell(OfficeXmlNode row, OfficeXmlNode cell, int index) {
            cell = row.Append(cell);
            cell.SetAttribute("t", "s");
            AppendValue(cell, index.ToString(CultureInfo.InvariantCulture));
        }

        private static void AppendValue(OfficeXmlNode cell, string value) {
            OfficeXmlNode element = cell.Append(cell.CreateElement(SpreadsheetML, "v"));
            element.SetInnerText(value);
        }

        // Numbers are written as the SDK cell value writes a double: the invariant culture's default format.
        private static string FormatNumber(double value) => value.ToString(CultureInfo.InvariantCulture);

        private static bool IsSimpleChartDataTable(SpreadsheetMarkup markup, OfficeXmlNode table) {
            if (table.IsNone) return false;
            List<OfficeXmlNode> columns = markup.Children(markup.Child(table, "tableColumns"), "tableColumn");
            return table.Attribute("ref") != null
                && columns.Count > 1
                && columns.All(column =>
                    markup.Child(column, "calculatedColumnFormula").IsNone
                    && markup.Child(column, "totalsRowFormula").IsNone)
                && markup.Descendants(table, "sortState").Count == 0;
        }

        private static string GetExcelColumn(int index) {
            string result = string.Empty;
            while (index > 0) {
                index--;
                result = (char)('A' + index % 26) + result;
                index /= 26;
            }
            return result;
        }
    }
}
