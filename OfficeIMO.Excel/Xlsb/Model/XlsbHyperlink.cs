namespace OfficeIMO.Excel.Xlsb.Model {
    internal sealed class XlsbHyperlink {
        internal XlsbHyperlink(
            XlsbCellRange range,
            string relationshipId,
            string? externalTarget,
            string location,
            string tooltip,
            string display) {
            Range = range;
            RelationshipId = relationshipId;
            ExternalTarget = externalTarget;
            Location = location;
            Tooltip = tooltip;
            Display = display;
        }

        internal XlsbCellRange Range { get; }

        internal string RelationshipId { get; }

        internal string? ExternalTarget { get; }

        internal string Location { get; }

        internal string Tooltip { get; }

        internal string Display { get; }

        internal bool IsExternal => ExternalTarget != null;
    }
}
