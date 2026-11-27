using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>A table definition part as the owned reader reads it: the attributes and children table reads use.</summary>
    internal sealed class ExcelTablePart {
        internal string? Name { get; set; }
        internal string? DisplayName { get; set; }
        internal string? Reference { get; set; }

        /// <summary>The header row count as written; null when the attribute is absent (one header row).</summary>
        internal uint? HeaderRowCount { get; set; }

        /// <summary>The totals row count as written; null when the attribute is absent.</summary>
        internal uint? TotalsRowCount { get; set; }

        internal bool TotalsRowShown { get; set; }
        internal bool HasAutoFilter { get; set; }
        internal bool HasStyleInfo { get; set; }
        internal string? StyleName { get; set; }
        internal bool ShowFirstColumn { get; set; }
        internal bool ShowLastColumn { get; set; }
        internal bool ShowRowStripes { get; set; }
        internal bool ShowColumnStripes { get; set; }
        internal List<(string? Name, string? TotalsRowFunction)> Columns { get; } = new();

        internal bool Matches(string nameOrDisplayName) =>
            string.Equals(Name, nameOrDisplayName, StringComparison.OrdinalIgnoreCase)
            || string.Equals(DisplayName, nameOrDisplayName, StringComparison.OrdinalIgnoreCase);

        /// <summary>The table as formulas resolve it; an absent totals row count reads as one, as the formula model has.</summary>
        internal ExcelFormulaTable ToFormulaTable() => new ExcelFormulaTable(Name, DisplayName, Reference, HeaderRowCount ?? 1U,
            TotalsRowCount ?? 1U, TotalsRowShown, Columns.ConvertAll(column => column.Name));

        /// <summary>The table as table listings describe it.</summary>
        internal ExcelTableInfo ToInfo(string sheetName, int sheetIndex) {
            string name = Name ?? DisplayName ?? string.Empty;
            var columns = new List<ExcelTableColumnInfo>(Columns.Count);
            for (int index = 0; index < Columns.Count; index++) {
                columns.Add(new ExcelTableColumnInfo(index + 1, Columns[index].Name ?? string.Empty, Columns[index].TotalsRowFunction));
            }
            return new ExcelTableInfo(name, DisplayName ?? name, Reference ?? string.Empty, sheetName, sheetIndex,
                HasStyleInfo ? StyleName : null, (HeaderRowCount ?? 1U) > 0U, (TotalsRowCount ?? 0U) > 0U, HasAutoFilter,
                ShowFirstColumn, ShowLastColumn, ShowRowStripes, ShowColumnStripes, columns);
        }

        /// <summary>Reads a table part's markup, or returns null when its root is not a table.</summary>
        internal static ExcelTablePart? Read(OfficeXmlPullReader reader) {
            ExcelTablePart? table = null;
            bool columnsSeen = false, inColumns = false, styleSeen = false, filterSeen = false;
            while (reader.Read()) {
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                if (reader.Depth == 1) {
                    if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "table")) return null;
                    table = new ExcelTablePart {
                        Name = Attribute(reader, "name"),
                        DisplayName = Attribute(reader, "displayName"),
                        Reference = Attribute(reader, "ref"),
                        HeaderRowCount = UInt(reader, "headerRowCount"),
                        TotalsRowCount = UInt(reader, "totalsRowCount"),
                        TotalsRowShown = IsTrue(reader, "totalsRowShown"),
                    };
                    continue;
                }
                if (table == null) return null;
                if (reader.Depth == 2) {
                    inColumns = false;
                    if (reader.NamespaceId != OfficeXmlNamespaces.SpreadsheetML) continue;
                    // The first element of each kind counts, as the SDK finds it.
                    if (reader.LocalNameEquals("tableColumns") && !columnsSeen) {
                        inColumns = columnsSeen = true;
                    } else if (reader.LocalNameEquals("autoFilter") && !filterSeen) {
                        filterSeen = table.HasAutoFilter = true;
                    } else if (reader.LocalNameEquals("tableStyleInfo") && !styleSeen) {
                        styleSeen = table.HasStyleInfo = true;
                        table.StyleName = Attribute(reader, "name");
                        table.ShowFirstColumn = IsTrue(reader, "showFirstColumn");
                        table.ShowLastColumn = IsTrue(reader, "showLastColumn");
                        table.ShowRowStripes = IsTrue(reader, "showRowStripes");
                        table.ShowColumnStripes = IsTrue(reader, "showColumnStripes");
                    }
                    continue;
                }
                if (reader.Depth == 3 && inColumns && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "tableColumn")) {
                    table.Columns.Add((Attribute(reader, "name"), Attribute(reader, "totalsRowFunction")));
                }
            }
            return table;
        }

        /// <summary>The tables of a worksheet part, in the order of its relationships.</summary>
        internal static List<ExcelTablePart> ReadSheetTables(OfficeOpcPackage package, string worksheetPartName) {
            var tables = new List<ExcelTablePart>();
            foreach (OfficeOpcRelationship relationship in package.GetRelationships(worksheetPartName)) {
                if (!string.Equals(relationship.Type, TableRelationshipType, StringComparison.Ordinal)
                    || relationship.IsExternal || relationship.TargetPartName is not { } target || !package.PartExists(target)) continue;
                OfficeXmlPullReader reader = package.OpenXml(target, out OfficePooledBytes content, processMarkupCompatibility: false);
                using (content) {
                    if (Read(reader) is { } table) tables.Add(table);
                }
            }
            return tables;
        }

        internal const string TableRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "table";

        private static string? Attribute(OfficeXmlPullReader reader, string name) =>
            reader.TryGetAttribute(OfficeXmlNamespaces.None, name, out string value) ? value : null;

        private static uint? UInt(OfficeXmlPullReader reader, string name) =>
            reader.TryGetAttribute(OfficeXmlNamespaces.None, name, out string value)
            && uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsed) ? parsed : null;

        private static bool IsTrue(OfficeXmlPullReader reader, string name) =>
            reader.TryGetAttribute(OfficeXmlNamespaces.None, name, out string value) && value is "1" or "true";
    }
}
