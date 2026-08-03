namespace OfficeIMO.PowerPoint {
    /// <summary>Identifies the physical PowerPoint file format used for loading or saving.</summary>
    public enum PowerPointFileFormat {
        /// <summary>Open XML presentation package (.pptx).</summary>
        Pptx = 0,

        /// <summary>PowerPoint 97-2003 binary presentation (.ppt).</summary>
        Ppt = 1,

        /// <summary>PowerPoint 97-2003 binary presentation template (.pot).</summary>
        Pot = 2,

        /// <summary>PowerPoint 97-2003 binary slide show (.pps).</summary>
        Pps = 3,

        /// <summary>Macro-enabled Open XML presentation package (.pptm).</summary>
        Pptm = 4
    }
}
