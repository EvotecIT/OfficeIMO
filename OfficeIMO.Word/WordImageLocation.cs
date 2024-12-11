using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    public class WordImageLocation {
        public ImagePart ImagePart { get; set; }
        public string RelationshipId { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public string ImageName { get; set; }
    }
}
