using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents an image placed on a slide.
    /// </summary>
    public class PPPicture : PPShape {
        internal PPPicture(Picture picture) : base(picture) {
        }
    }
}

