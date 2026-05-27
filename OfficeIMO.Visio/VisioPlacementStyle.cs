namespace OfficeIMO.Visio {
    /// <summary>
    /// Specifies the page-level Visio placement style used when laying out shapes.
    /// </summary>
    public enum VisioPlacementStyle {
        /// <summary>Use Visio's default placement style.</summary>
        Default = 0,

        /// <summary>Place flow/tree shapes from top to bottom.</summary>
        TopToBottom = 1,

        /// <summary>Place flow/tree shapes from left to right.</summary>
        LeftToRight = 2,

        /// <summary>Place flow/tree shapes radially.</summary>
        Radial = 3,

        /// <summary>Place flow/tree shapes from bottom to top.</summary>
        BottomToTop = 4,

        /// <summary>Place flow/tree shapes from right to left.</summary>
        RightToLeft = 5,

        /// <summary>Place shapes in a circle.</summary>
        Circular = 6,

        /// <summary>Use compact tree placement downward, then right.</summary>
        CompactDownRight = 7,

        /// <summary>Use compact tree placement rightward, then down.</summary>
        CompactRightDown = 8,

        /// <summary>Use compact tree placement rightward, then up.</summary>
        CompactRightUp = 9,

        /// <summary>Use compact tree placement upward, then right.</summary>
        CompactUpRight = 10,

        /// <summary>Use compact tree placement upward, then left.</summary>
        CompactUpLeft = 11,

        /// <summary>Use compact tree placement leftward, then up.</summary>
        CompactLeftUp = 12,

        /// <summary>Use compact tree placement leftward, then down.</summary>
        CompactLeftDown = 13,

        /// <summary>Use compact tree placement downward, then left.</summary>
        CompactDownLeft = 14,

        /// <summary>Use parent default placement.</summary>
        ParentDefault = 15,

        /// <summary>Place hierarchical shapes top-to-bottom, left aligned.</summary>
        HierarchyTopToBottomLeft = 16,

        /// <summary>Place hierarchical shapes top-to-bottom, centered.</summary>
        HierarchyTopToBottomCenter = 17,

        /// <summary>Place hierarchical shapes top-to-bottom, right aligned.</summary>
        HierarchyTopToBottomRight = 18,

        /// <summary>Place hierarchical shapes bottom-to-top, left aligned.</summary>
        HierarchyBottomToTopLeft = 19,

        /// <summary>Place hierarchical shapes bottom-to-top, centered.</summary>
        HierarchyBottomToTopCenter = 20,

        /// <summary>Place hierarchical shapes bottom-to-top, right aligned.</summary>
        HierarchyBottomToTopRight = 21,

        /// <summary>Place hierarchical shapes left-to-right, top aligned.</summary>
        HierarchyLeftToRightTop = 22,

        /// <summary>Place hierarchical shapes left-to-right, middle aligned.</summary>
        HierarchyLeftToRightMiddle = 23,

        /// <summary>Place hierarchical shapes left-to-right, bottom aligned.</summary>
        HierarchyLeftToRightBottom = 24,

        /// <summary>Place hierarchical shapes right-to-left, top aligned.</summary>
        HierarchyRightToLeftTop = 25,

        /// <summary>Place hierarchical shapes right-to-left, middle aligned.</summary>
        HierarchyRightToLeftMiddle = 26,

        /// <summary>Place hierarchical shapes right-to-left, bottom aligned.</summary>
        HierarchyRightToLeftBottom = 27
    }
}
