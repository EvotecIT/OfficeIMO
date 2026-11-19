using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Xlsb.Model {
    /// <summary>A worksheet table decoded from a native Table part (MS-XLSB 2.1.7.51).</summary>
    internal sealed class XlsbTable {
        internal XlsbTable(string relationshipId, string partName, Table definition, uint id, string displayName,
            XlsbCellRange range, uint headerRows, uint totalsRows, IReadOnlyList<string> columns, bool hasUnsupportedContent) {
            RelationshipId = relationshipId;
            PartName = partName;
            Definition = definition;
            Id = id;
            DisplayName = displayName;
            Range = range;
            HeaderRows = headerRows;
            TotalsRows = totalsRows;
            Columns = columns;
            HasUnsupportedContent = hasUnsupportedContent;
        }

        /// <summary>The worksheet relationship id that BrtListPart names.</summary>
        internal string RelationshipId { get; }

        /// <summary>The Table part's package path without a leading slash.</summary>
        internal string PartName { get; }

        /// <summary>The canonical SpreadsheetML projection; clone it before attaching it to a document.</summary>
        internal Table Definition { get; }

        internal uint Id { get; }

        internal string DisplayName { get; }

        internal XlsbCellRange Range { get; }

        internal uint HeaderRows { get; }

        internal uint TotalsRows { get; }

        /// <summary>The column captions in table order.</summary>
        internal IReadOnlyList<string> Columns { get; }

        /// <summary>True when the part holds records the projection does not represent (sort state, XML maps, extensions).</summary>
        internal bool HasUnsupportedContent { get; set; }

        /// <summary>
        /// False when the worksheet keeps its tables native because another of its table parts is unreadable. Formulas that
        /// name such a table are not decoded either, so no projected name or cell refers to a table missing from the model.
        /// </summary>
        internal bool IsProjected { get; set; } = true;
    }
}
