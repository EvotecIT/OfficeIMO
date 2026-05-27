namespace OfficeIMO.Visio {
    /// <summary>
    /// Specifies Visio's page-level connector routing style.
    /// </summary>
    public enum VisioPageRouteStyle {
        /// <summary>Use Visio's default right-angle routing.</summary>
        Default = 0,

        /// <summary>Use right-angle routing.</summary>
        RightAngle = 1,

        /// <summary>Use straight routing.</summary>
        Straight = 2,

        /// <summary>Use organization-chart routing from top to bottom.</summary>
        OrganizationChartTopToBottom = 3,

        /// <summary>Use organization-chart routing from left to right.</summary>
        OrganizationChartLeftToRight = 4,

        /// <summary>Use flowchart routing from top to bottom.</summary>
        FlowchartTopToBottom = 5,

        /// <summary>Use flowchart routing from left to right.</summary>
        FlowchartLeftToRight = 6,

        /// <summary>Use tree routing from top to bottom.</summary>
        TreeTopToBottom = 7,

        /// <summary>Use tree routing from left to right.</summary>
        TreeLeftToRight = 8,

        /// <summary>Use network routing.</summary>
        Network = 9,

        /// <summary>Use organization-chart routing from bottom to top.</summary>
        OrganizationChartBottomToTop = 10,

        /// <summary>Use organization-chart routing from right to left.</summary>
        OrganizationChartRightToLeft = 11,

        /// <summary>Use flowchart routing from bottom to top.</summary>
        FlowchartBottomToTop = 12,

        /// <summary>Use flowchart routing from right to left.</summary>
        FlowchartRightToLeft = 13,

        /// <summary>Use tree routing from bottom to top.</summary>
        TreeBottomToTop = 14,

        /// <summary>Use tree routing from right to left.</summary>
        TreeRightToLeft = 15,

        /// <summary>Route from center to center.</summary>
        CenterToCenter = 16,

        /// <summary>Use simple routing from top to bottom.</summary>
        SimpleTopToBottom = 17,

        /// <summary>Use simple routing from left to right.</summary>
        SimpleLeftToRight = 18,

        /// <summary>Use simple routing from bottom to top.</summary>
        SimpleBottomToTop = 19,

        /// <summary>Use simple routing from right to left.</summary>
        SimpleRightToLeft = 20,

        /// <summary>Use simple horizontal-then-vertical routing.</summary>
        SimpleHorizontalVertical = 21,

        /// <summary>Use simple vertical-then-horizontal routing.</summary>
        SimpleVerticalHorizontal = 22
    }
}
