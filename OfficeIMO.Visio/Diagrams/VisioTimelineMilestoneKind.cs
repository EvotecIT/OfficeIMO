namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Semantic milestone kinds used by the timeline diagram builder.
    /// </summary>
    public enum VisioTimelineMilestoneKind {
        /// <summary>A standard milestone.</summary>
        Milestone,

        /// <summary>A release or delivery marker.</summary>
        Release,

        /// <summary>A decision or approval marker.</summary>
        Decision,

        /// <summary>A risk, issue, or attention marker.</summary>
        Risk
    }
}
