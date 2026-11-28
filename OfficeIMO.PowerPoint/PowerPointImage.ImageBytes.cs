using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointImage {
        private const string SvgBlipNamespace = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
        private OfficeXmlNode PictureBlipFill => XmlElement.Child(OfficeXmlNamespaces.PresentationML, "blipFill");
        private OfficeXmlNode ImageBlip => PictureBlipFill.Child(OfficeXmlNamespaces.DrawingML, "blip");

        /// <summary>Returns a snapshot of the embedded image bytes for export adapters.</summary>
        public byte[] GetImageBytes() => GetImageBytes(maximumBytes: null);

        /// <summary>Returns the raster fallback of an SVG picture, for renderers that cannot draw its SVG.</summary>
        internal bool TryGetRasterFallback(int? maximumBytes, out byte[] bytes) {
            PowerPointPartRef fallback = GetSvgRelationshipId() == null ? default : GetFallbackImagePartRef();
            bytes = fallback.IsNone ? Array.Empty<byte>() : fallback.ReadBytes(maximumBytes);
            return bytes.Length > 0;
        }

        /// <summary>Returns a bounded snapshot of the embedded image bytes.</summary>
        internal byte[] GetImageBytes(int? maximumBytes) {
            PowerPointPartRef image = GetImagePartRef();
            if (image.IsNone) throw new InvalidOperationException("Picture has no embedded image part.");
            return image.ReadBytes(maximumBytes);
        }

        // A picture with an SVG source draws that source. Its ordinary blip is the raster fallback.
        private PowerPointPartRef GetImagePartRef() {
            PowerPointPartRef svg = ResolveImagePart(GetSvgRelationshipId());
            return svg.IsNone ? GetFallbackImagePartRef() : svg;
        }

        private PowerPointPartRef GetFallbackImagePartRef() =>
            ResolveImagePart(ImageBlip.Attribute(OfficeXmlNamespaces.Relationships, "embed"));

        private string? GetSvgRelationshipId() {
            const int drawing = OfficeXmlNamespaces.DrawingML;
            foreach (OfficeXmlNode extension in ImageBlip.Child(drawing, "extLst").Children(drawing, "ext")) {
                string? id = extension.Child(SvgBlipNamespace, "svgBlip").Attribute(OfficeXmlNamespaces.Relationships, "embed");
                if (!string.IsNullOrEmpty(id)) return id;
            }
            return null;
        }

        private PowerPointPartRef ResolveImagePart(string? relationshipId) =>
            relationshipId != null && OwnerPartRef.TryGetRelationship(relationshipId, out PowerPointRelationship relationship)
                && OfficeXmlNamespaces.NormalizeRelationshipType(relationship.Type) == OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "image"
                ? relationship.TargetPart : default;
    }
}
