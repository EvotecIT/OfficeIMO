namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a Visio master.
    /// </summary>
    public class VisioMaster {
        /// <summary>
        /// Initializes a new instance of the <see cref="VisioMaster"/> class.
        /// </summary>
        /// <param name="id">Identifier of the master.</param>
        /// <param name="nameU">Universal name of the master.</param>
        /// <param name="shape">Associated master shape.</param>
        public VisioMaster(string id, string nameU, VisioShape shape) {
            Id = id;
            NameU = nameU;
            Shape = shape;
        }

        /// <summary>
        /// Gets the master identifier.
        /// </summary>
        public string Id { get; }

        /// <summary>
        /// Gets the universal name of the master.
        /// </summary>
        public string NameU { get; }

        /// <summary>
        /// Gets the shape that defines the master.
        /// </summary>
        public VisioShape Shape { get; }

        /// <summary>
        /// Optional canonical MasterContents XML to write verbatim (used when importing from a template VSDX).
        /// When set, the save routine will output this document instead of generating geometry from <see cref="Shape"/>.
        /// </summary>
        public System.Xml.Linq.XDocument? TemplateXml { get; set; }

        /// <summary>
        /// Optional canonical Master element from masters.xml (attributes, PageSheet, Icon, etc.).
        /// When provided, the save routine will base the masters.xml entry on this element to match Visio assets.
        /// </summary>
        public System.Xml.Linq.XElement? TemplateMasterElement { get; set; }
    }
}
