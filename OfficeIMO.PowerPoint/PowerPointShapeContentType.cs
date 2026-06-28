namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Describes the primary content carried by a PowerPoint shape wrapper.
    /// </summary>
    public enum PowerPointShapeContentType {
        /// <summary>
        ///     The shape content type could not be identified.
        /// </summary>
        Unknown,

        /// <summary>
        ///     The shape is a regular drawing shape.
        /// </summary>
        AutoShape,

        /// <summary>
        ///     The shape is a text box.
        /// </summary>
        TextBox,

        /// <summary>
        ///     The shape is a picture.
        /// </summary>
        Picture,

        /// <summary>
        ///     The shape is a table.
        /// </summary>
        Table,

        /// <summary>
        ///     The shape is a chart.
        /// </summary>
        Chart,

        /// <summary>
        ///     The shape is a group of shapes.
        /// </summary>
        Group,

        /// <summary>
        ///     The shape represents embedded media, such as audio or video.
        /// </summary>
        Media,

        /// <summary>
        ///     The shape represents a SmartArt diagram.
        /// </summary>
        SmartArt,

        /// <summary>
        ///     The shape represents an embedded OLE object.
        /// </summary>
        OleObject,

        /// <summary>
        ///     The shape is a native PowerPoint connector.
        /// </summary>
        Connector
    }
}
