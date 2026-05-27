namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Semantic node kinds used by the org chart diagram builder.
    /// </summary>
    public enum VisioOrgChartNodeKind {
        /// <summary>The top executive/root node.</summary>
        Executive,

        /// <summary>A manager or team lead.</summary>
        Manager,

        /// <summary>A standard reporting position.</summary>
        Position,

        /// <summary>An assistant attached beside a manager.</summary>
        Assistant,

        /// <summary>An open position.</summary>
        Vacancy,

        /// <summary>An external advisor, vendor, or partner role.</summary>
        External
    }
}
