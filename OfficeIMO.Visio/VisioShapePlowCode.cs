namespace OfficeIMO.Visio {
    /// <summary>
    /// Specifies how a shape moves away when another placeable shape is dropped nearby.
    /// </summary>
    public enum VisioShapePlowCode {
        /// <summary>Use the page-level plow setting.</summary>
        PageDefault = 0,

        /// <summary>Do not move shapes away.</summary>
        Never = 1,

        /// <summary>Move every affected shape away.</summary>
        Always = 2
    }
}
