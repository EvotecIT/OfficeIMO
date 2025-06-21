using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
        public OpenXmlPartContainer PartContainer { get; set; }
    /// <summary>
    /// Represents the location and metadata of an image within a document.
    /// </summary>
    public class WordImageLocation {
        /// <summary>
        /// Gets or sets the <see cref="ImagePart"/> associated with the image.
        /// </summary>
        public ImagePart ImagePart { get; set; }

        /// <summary>
        /// Gets or sets the relationship identifier linking to the image part.
        /// </summary>
        public string RelationshipId { get; set; }

        /// <summary>
        /// Gets or sets the width of the image in pixels.
        /// </summary>
        public double Width { get; set; }

        /// <summary>
        /// Gets or sets the height of the image in pixels.
        /// </summary>
        public double Height { get; set; }

        /// <summary>
        /// Gets or sets the descriptive name of the image.
        /// </summary>
        public string ImageName { get; set; }
    }
}
