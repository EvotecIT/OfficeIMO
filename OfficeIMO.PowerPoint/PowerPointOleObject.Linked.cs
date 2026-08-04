using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointOleObject {
        /// <summary>Gets whether this OLE frame references an external object.</summary>
        public bool IsLinked => OleObject.GetFirstChild<P.OleObjectLink>() != null;

        /// <summary>Gets whether this OLE frame owns a package payload.</summary>
        public bool IsEmbedded => OleObject.GetFirstChild<P.OleObjectEmbed>() != null;

        /// <summary>Gets the external object URI when linked.</summary>
        public Uri? LinkUri {
            get {
                if (!IsLinked) return null;
                string? id = OleObject.Id?.Value;
                return _slidePart.ExternalRelationships.FirstOrDefault(item => item.Id == id)?.Uri;
            }
        }

        /// <summary>Gets or sets whether PowerPoint automatically refreshes the linked object.</summary>
        public bool AutoUpdate {
            get => OleObject.GetFirstChild<P.OleObjectLink>()?.AutoUpdate?.Value == true;
            set {
                P.OleObjectLink link = OleObject.GetFirstChild<P.OleObjectLink>()
                    ?? throw new InvalidOperationException("The OLE object is embedded rather than linked.");
                link.AutoUpdate = value;
            }
        }

        /// <summary>Gets the embedded MIME type, or null for linked objects.</summary>
        public string? EmbeddedContentType => IsEmbedded ? EmbeddedPart.ContentType : null;

        /// <summary>Updates the external target without replacing preview or OLE metadata.</summary>
        public PowerPointOleObject UpdateLink(Uri uri) {
            if (uri == null) throw new ArgumentNullException(nameof(uri));
            if (!uri.IsAbsoluteUri) throw new ArgumentException("A linked OLE URI must be absolute.", nameof(uri));
            string id = OleObject.Id?.Value ?? throw new InvalidOperationException("The OLE object has no relationship.");
            ExternalRelationship relationship = _slidePart.ExternalRelationships
                .FirstOrDefault(item => item.Id == id)
                ?? throw new InvalidOperationException("The OLE object is embedded rather than linked.");
            _slidePart.DeleteExternalRelationship(id);
            _slidePart.AddExternalRelationship(relationship.RelationshipType, uri, id);
            return this;
        }
    }
}
