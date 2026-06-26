namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies OfficeArt record types found in legacy XLS MsoDrawing payload headers.
    /// </summary>
    public enum LegacyXlsDrawingEscherRecordType {
        /// <summary>Container for document-wide OfficeArt records.</summary>
        OfficeArtDggContainer = 0xF000,

        /// <summary>Container for BLIP store entries.</summary>
        OfficeArtBStoreContainer = 0xF001,

        /// <summary>Container for drawing-wide OfficeArt records.</summary>
        OfficeArtDgContainer = 0xF002,

        /// <summary>Container for grouped shapes.</summary>
        OfficeArtSpgrContainer = 0xF003,

        /// <summary>Container for a shape.</summary>
        OfficeArtSpContainer = 0xF004,

        /// <summary>Container for drawing solver rules.</summary>
        OfficeArtSolverContainer = 0xF005,

        /// <summary>Document-wide drawing-group block.</summary>
        OfficeArtFDGGBlock = 0xF006,

        /// <summary>BLIP store entry.</summary>
        OfficeArtFBSE = 0xF007,

        /// <summary>Drawing data record.</summary>
        OfficeArtFDG = 0xF008,

        /// <summary>Shape group record.</summary>
        OfficeArtFSPGR = 0xF009,

        /// <summary>Shape record.</summary>
        OfficeArtFSP = 0xF00A,

        /// <summary>Shape properties record.</summary>
        OfficeArtFOPT = 0xF00B,

        /// <summary>Client text box record.</summary>
        OfficeArtFClientTextbox = 0xF00D,

        /// <summary>Shape child-anchor record.</summary>
        OfficeArtChildAnchor = 0xF00F,

        /// <summary>Client-anchor record.</summary>
        OfficeArtFClientAnchor = 0xF010,

        /// <summary>Client-data record.</summary>
        OfficeArtFClientData = 0xF011,

        /// <summary>Container for split-menu color MRU entries.</summary>
        OfficeArtSplitMenuColorContainer = 0xF11E
    }
}
