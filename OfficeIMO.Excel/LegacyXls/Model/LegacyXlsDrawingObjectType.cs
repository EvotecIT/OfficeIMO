namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies the object type declared by the FtCmo structure in a legacy XLS Obj record.
    /// </summary>
    public enum LegacyXlsDrawingObjectType {
        /// <summary>Group object.</summary>
        Group = 0x0000,

        /// <summary>Line object.</summary>
        Line = 0x0001,

        /// <summary>Rectangle object.</summary>
        Rectangle = 0x0002,

        /// <summary>Oval object.</summary>
        Oval = 0x0003,

        /// <summary>Arc object.</summary>
        Arc = 0x0004,

        /// <summary>Chart object.</summary>
        Chart = 0x0005,

        /// <summary>Text object.</summary>
        Text = 0x0006,

        /// <summary>Button object.</summary>
        Button = 0x0007,

        /// <summary>Picture object.</summary>
        Picture = 0x0008,

        /// <summary>Polygon object.</summary>
        Polygon = 0x0009,

        /// <summary>Checkbox object.</summary>
        Checkbox = 0x000B,

        /// <summary>Radio button object.</summary>
        RadioButton = 0x000C,

        /// <summary>Edit box object.</summary>
        EditBox = 0x000D,

        /// <summary>Label object.</summary>
        Label = 0x000E,

        /// <summary>Dialog box object.</summary>
        DialogBox = 0x000F,

        /// <summary>Spin control object.</summary>
        SpinControl = 0x0010,

        /// <summary>Scrollbar object.</summary>
        Scrollbar = 0x0011,

        /// <summary>List object.</summary>
        List = 0x0012,

        /// <summary>Group box object.</summary>
        GroupBox = 0x0013,

        /// <summary>Dropdown list object.</summary>
        DropdownList = 0x0014,

        /// <summary>Note object.</summary>
        Note = 0x0019,

        /// <summary>OfficeArt object.</summary>
        OfficeArtObject = 0x001E
    }
}
