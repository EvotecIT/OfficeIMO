using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// What the workbook's cell metadata says about each cell metadata index (<c>cm</c>): whether it marks a dynamic
    /// array, and whether that array is collapsed to one cell.
    /// </summary>
    internal sealed class ExcelCellMetadataIndex {
        private const string DynamicArrayType = "XLDAPR";
        private readonly List<(bool Dynamic, bool Collapsed)?> _cells;

        private ExcelCellMetadataIndex(List<(bool Dynamic, bool Collapsed)?> cells, bool hasDynamicArrayType) {
            _cells = cells;
            HasDynamicArrayType = hasDynamicArrayType;
        }

        /// <summary>Whether the metadata declares the dynamic array type at all.</summary>
        internal bool HasDynamicArrayType { get; }

        /// <summary>The dynamic array properties a one-based cell metadata index names, or null when it names none.</summary>
        internal (bool Dynamic, bool Collapsed)? DynamicArray(uint oneBasedIndex) =>
            oneBasedIndex == 0 || oneBasedIndex > _cells.Count ? null : _cells[(int)oneBasedIndex - 1];

        /// <summary>The array formula info of a cell, or null when the cell holds no array formula with a range.</summary>
        internal static ExcelFormulaArrayInfo? ArrayInfo(ExcelFormulaCell cell, ExcelCellMetadataIndex? metadata) {
            if (cell.FormulaKind != ExcelFormulaKind.Array || string.IsNullOrWhiteSpace(cell.FormulaReference)) return null;
            uint? metadataIndex = cell.CellMetaIndex;
            (bool Dynamic, bool Collapsed)? properties = metadataIndex is uint index && index > 0 ? metadata?.DynamicArray(index) : null;
            bool dynamic = properties?.Dynamic == true;
            bool collapsed = dynamic && properties!.Value.Collapsed;
            return new ExcelFormulaArrayInfo(cell.FormulaReference!, dynamic, collapsed, metadataIndex);
        }

        /// <summary>Reads the metadata part's markup.</summary>
        internal static ExcelCellMetadataIndex Read(byte[] xml, int length) {
            var reader = new OfficeXmlPullReader(xml, length, processMarkupCompatibility: false);
            var typeNames = new List<string?>();
            var cellRecords = new List<List<(uint Type, uint Value)>>();
            // The first future metadata block list for each name, as a lookup by name finds it.
            var future = new Dictionary<string, List<(bool Dynamic, bool Collapsed)?>>(StringComparer.OrdinalIgnoreCase);
            List<(bool Dynamic, bool Collapsed)?>? currentFuture = null;
            List<(uint Type, uint Value)>? currentBlock = null;
            int section = 0; // 1 types, 2 cell metadata, 3 future metadata
            while (reader.Read()) {
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                int depth = reader.Depth;
                if (depth == 2) {
                    section = 0;
                    currentFuture = null;
                    if (reader.NamespaceId != OfficeXmlNamespaces.SpreadsheetML) continue;
                    if (reader.LocalNameEquals("metadataTypes")) section = 1;
                    else if (reader.LocalNameEquals("cellMetadata")) section = 2;
                    else if (reader.LocalNameEquals("futureMetadata")) {
                        section = 3;
                        string name = reader.TryGetAttribute(OfficeXmlNamespaces.None, "name", out string futureName) ? futureName : string.Empty;
                        if (!future.ContainsKey(name)) future.Add(name, currentFuture = new List<(bool Dynamic, bool Collapsed)?>());
                    }
                    continue;
                }
                if (section == 1 && depth == 3 && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "metadataType")) {
                    typeNames.Add(reader.TryGetAttribute(OfficeXmlNamespaces.None, "name", out string typeName) ? typeName : null);
                } else if (section == 2 && depth == 3 && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "bk")) {
                    cellRecords.Add(currentBlock = new List<(uint Type, uint Value)>());
                } else if (section == 2 && depth == 4 && currentBlock != null && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "rc")) {
                    uint type = ReadUInt(reader, "t") ?? 0U;
                    uint value = ReadUInt(reader, "v") ?? uint.MaxValue;
                    currentBlock.Add((type, value));
                } else if (section == 3 && depth == 3 && currentFuture != null && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "bk")) {
                    currentFuture.Add(ReadDynamicArrayProperties(reader));
                }
            }

            bool hasDynamicType = typeNames.Exists(name => string.Equals(name, DynamicArrayType, StringComparison.OrdinalIgnoreCase));
            var cells = new List<(bool Dynamic, bool Collapsed)?>(cellRecords.Count);
            foreach (List<(uint Type, uint Value)> block in cellRecords) {
                (bool Dynamic, bool Collapsed)? found = null;
                foreach ((uint type, uint value) in block) {
                    if (type == 0 || type > typeNames.Count) continue;
                    string? typeName = typeNames[(int)type - 1];
                    if (!string.Equals(typeName, DynamicArrayType, StringComparison.OrdinalIgnoreCase)) continue;
                    if (!future.TryGetValue(typeName!, out List<(bool Dynamic, bool Collapsed)?>? blocks)
                        || value > int.MaxValue || value >= blocks.Count || blocks[(int)value] is not { } properties) continue;
                    found = properties;
                    break;
                }
                cells.Add(found);
            }
            return new ExcelCellMetadataIndex(cells, hasDynamicType);
        }

        private static (bool Dynamic, bool Collapsed)? ReadDynamicArrayProperties(OfficeXmlPullReader reader) {
            if (reader.IsEmptyElement) return null;
            int depth = reader.Depth;
            (bool Dynamic, bool Collapsed)? found = null;
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) break;
                if (found != null || reader.NodeKind != OfficeXmlNodeKind.StartElement || !reader.LocalNameEquals("dynamicArrayProperties")) continue;
                string? dynamicValue = FindAttribute(reader, "fDynamic");
                string? collapsedValue = FindAttribute(reader, "fCollapsed");
                bool dynamic = !string.Equals(dynamicValue, "0", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(dynamicValue, "false", StringComparison.OrdinalIgnoreCase);
                bool collapsed = dynamic && (string.Equals(collapsedValue, "1", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(collapsedValue, "true", StringComparison.OrdinalIgnoreCase));
                found = (dynamic, collapsed);
            }
            return found;
        }

        // Matches the attribute by local name in any namespace, as the metadata reader always has.
        private static string? FindAttribute(OfficeXmlPullReader reader, string localName) {
            for (int index = 0; index < reader.AttributeCount; index++) {
                if (string.Equals(reader.GetAttributeLocalName(index), localName, StringComparison.OrdinalIgnoreCase)) {
                    return reader.GetAttributeValue(index);
                }
            }
            return null;
        }

        private static uint? ReadUInt(OfficeXmlPullReader reader, string attribute) =>
            reader.TryGetAttribute(OfficeXmlNamespaces.None, attribute, out string text)
            && uint.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out uint value) ? value : null;
    }
}
