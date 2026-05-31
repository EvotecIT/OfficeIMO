using System;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Describes preview/icon image metadata discovered for a stencil master.
    /// </summary>
    public sealed class VisioStencilPreviewImage {
        /// <summary>
        /// Initializes preview image metadata.
        /// </summary>
        public VisioStencilPreviewImage(string relationshipId, string target, string? contentType = null, string? extension = null, long? byteLength = null, bool isExternal = false) {
            if (string.IsNullOrWhiteSpace(relationshipId)) throw new ArgumentException("Relationship id cannot be null or whitespace.", nameof(relationshipId));
            if (string.IsNullOrWhiteSpace(target)) throw new ArgumentException("Preview image target cannot be null or whitespace.", nameof(target));

            RelationshipId = relationshipId;
            Target = target;
            ContentType = string.IsNullOrWhiteSpace(contentType) ? null : contentType;
            Extension = string.IsNullOrWhiteSpace(extension) ? null : extension;
            ByteLength = byteLength;
            IsExternal = isExternal;
        }

        /// <summary>Relationship id from the source master part.</summary>
        public string RelationshipId { get; }

        /// <summary>Original relationship target from the source package.</summary>
        public string Target { get; }

        /// <summary>Content type for the preview image, when known.</summary>
        public string? ContentType { get; }

        /// <summary>File extension for the preview image, when known.</summary>
        public string? Extension { get; }

        /// <summary>Byte length for embedded preview image payloads, when known.</summary>
        public long? ByteLength { get; }

        /// <summary>Whether the preview image relationship points outside the package.</summary>
        public bool IsExternal { get; }
    }
}
