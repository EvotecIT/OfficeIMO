namespace OfficeIMO.Excel.Xlsb.Model {
    internal sealed class XlsbWorksheet {
        private readonly List<XlsbCell> _cells = new List<XlsbCell>();

        internal XlsbWorksheet(string name, string relationshipId, uint tabId, uint state) {
            Name = name;
            RelationshipId = relationshipId;
            TabId = tabId;
            State = state;
        }

        internal string Name { get; }

        internal string RelationshipId { get; }

        internal uint TabId { get; }

        internal uint State { get; }

        internal string? PartName { get; set; }

        internal IReadOnlyList<XlsbCell> Cells => _cells;

        internal void AddCell(XlsbCell cell) => _cells.Add(cell);
    }
}
