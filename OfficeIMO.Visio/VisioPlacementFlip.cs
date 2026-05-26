using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Specifies how Visio may flip or rotate shapes during page layout.
    /// </summary>
    [Flags]
    public enum VisioPlacementFlip {
        /// <summary>Use Visio's default flip behavior.</summary>
        Default = 0,

        /// <summary>Allow horizontal flips.</summary>
        Horizontal = 1,

        /// <summary>Allow vertical flips.</summary>
        Vertical = 2,

        /// <summary>Allow 90-degree rotations.</summary>
        Rotate90 = 4,

        /// <summary>Do not flip shapes.</summary>
        None = 8
    }
}
