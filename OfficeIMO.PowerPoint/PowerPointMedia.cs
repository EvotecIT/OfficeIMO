using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents an embedded audio or video media frame on a slide.
    /// </summary>
    public class PowerPointMedia : PowerPointPicture {
        internal PowerPointMedia(Picture picture, SlidePart slidePart, PowerPointMediaKind kind) : base(picture, slidePart) {
            Kind = kind;
        }

        /// <summary>
        ///     Gets the media kind represented by this shape.
        /// </summary>
        public PowerPointMediaKind Kind { get; }

        /// <summary>
        ///     Gets the MIME content type of the embedded media part.
        /// </summary>
        public string? MediaContentType {
            get {
                MediaDataPart? mediaPart = GetMediaDataPart();
                return mediaPart?.ContentType;
            }
        }

        /// <summary>
        ///     Gets the relationship id used by the audio/video file reference.
        /// </summary>
        public string? MediaReferenceId {
            get {
                Picture picture = (Picture)Element;
                ApplicationNonVisualDrawingProperties? properties =
                    picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
                return Kind == PowerPointMediaKind.Audio
                    ? properties?.GetFirstChild<A.AudioFromFile>()?.Link?.Value
                    : properties?.GetFirstChild<A.VideoFromFile>()?.Link?.Value;
            }
        }

        /// <summary>
        ///     Gets the p14 media relationship id used by PowerPoint playback metadata.
        /// </summary>
        public string? PlaybackReferenceId {
            get {
                Picture picture = (Picture)Element;
                return picture.NonVisualPictureProperties?
                    .ApplicationNonVisualDrawingProperties?
                    .Descendants<P14.Media>()
                    .FirstOrDefault()?
                    .Embed?
                    .Value;
            }
        }

        /// <summary>Returns the exact embedded audio or video payload.</summary>
        public byte[] GetData() {
            MediaDataPart mediaPart = GetMediaDataPart()
                ?? throw new InvalidOperationException(
                    "The media shape has no embedded data relationship.");
            using Stream stream = mediaPart.GetStream(FileMode.Open,
                FileAccess.Read);
            return OfficeStreamReader.ReadAllBytes(stream);
        }

        /// <summary>Replaces the embedded audio or video payload.</summary>
        public void UpdateData(Stream media) {
            if (media == null) throw new ArgumentNullException(nameof(media));
            if (!media.CanRead) {
                throw new ArgumentException("Media stream must be readable.",
                    nameof(media));
            }
            MediaDataPart mediaPart = GetMediaDataPart()
                ?? throw new InvalidOperationException(
                    "The media shape has no embedded data relationship.");
            if (media.CanSeek) media.Position = 0;
            byte[] replacement = OfficeStreamReader.ReadAllBytes(media);
            byte[] original;
            using (Stream current = mediaPart.GetStream(FileMode.Open,
                       FileAccess.Read)) {
                original = OfficeStreamReader.ReadAllBytes(current);
            }
            if (IsSharedOutsideThisFrame(mediaPart)) {
                ReplaceSharedDataPart(mediaPart, replacement);
                return;
            }
            try {
                using var input = new MemoryStream(replacement,
                    writable: false);
                mediaPart.FeedData(input);
            } catch {
                using var rollback = new MemoryStream(original,
                    writable: false);
                mediaPart.FeedData(rollback);
                throw;
            }
        }

        private bool IsSharedOutsideThisFrame(MediaDataPart mediaPart) {
            SlidePart owner = SlidePart;
            if (owner.OpenXmlPackage is not PresentationDocument document) {
                return true;
            }
            var localRelationshipIds = new HashSet<string>(new[] {
                MediaReferenceId, PlaybackReferenceId
            }.Where(id => !string.IsNullOrEmpty(id)).Cast<string>(),
                StringComparer.Ordinal);
            foreach (string relationshipId in localRelationshipIds) {
                if (CountRelationshipReferences(owner.RootElement,
                        relationshipId)
                    > CountRelationshipReferences(Element,
                        relationshipId)) {
                    return true;
                }
            }
            foreach (SlidePart slidePart in document.PresentationPart?
                         .SlideParts ?? Enumerable.Empty<SlidePart>()) {
                foreach (DataPartReferenceRelationship relationship in
                         slidePart.DataPartReferenceRelationships) {
                    if (!ReferenceEquals(relationship.DataPart, mediaPart)) {
                        continue;
                    }
                    if (!ReferenceEquals(slidePart, owner)
                        || !localRelationshipIds.Contains(relationship.Id)) {
                        return true;
                    }
                }
            }
            return false;
        }

        private static int CountRelationshipReferences(
            OpenXmlElement? root, string relationshipId) {
            if (root == null) return 0;
            int count = root.GetAttributes().Count(attribute =>
                string.Equals(attribute.Value, relationshipId,
                    StringComparison.Ordinal));
            return count + root.Descendants().Sum(element => element
                .GetAttributes().Count(attribute => string.Equals(
                    attribute.Value, relationshipId,
                    StringComparison.Ordinal)));
        }

        private void ReplaceSharedDataPart(MediaDataPart original,
            byte[] replacement) {
            SlidePart owner = SlidePart;
            if (owner.OpenXmlPackage is not PresentationDocument document) {
                throw new InvalidOperationException(
                    "The media shape is not attached to a presentation document.");
            }
            string[] localRelationshipIds = new[] {
                MediaReferenceId, PlaybackReferenceId
            }.Where(id => !string.IsNullOrEmpty(id)).Cast<string>()
                .Distinct(StringComparer.Ordinal).ToArray();
            DataPartReferenceRelationship[] relationships = owner
                .DataPartReferenceRelationships.Where(relationship =>
                    localRelationshipIds.Contains(relationship.Id,
                        StringComparer.Ordinal)
                    && ReferenceEquals(relationship.DataPart, original))
                .ToArray();
            if (relationships.Length == 0) {
                throw new InvalidOperationException(
                    "The media shape has no replaceable data relationships.");
            }
            string extension = Path.GetExtension(original.Uri.OriginalString)
                .TrimStart('.');
            MediaDataPart detached = extension.Length == 0
                ? document.CreateMediaDataPart(original.ContentType)
                : document.CreateMediaDataPart(original.ContentType,
                    extension);
            var replacements = new List<(DataPartReferenceRelationship
                Source, string TargetId)>();
            var deletedOriginals = new List<
                DataPartReferenceRelationship>();
            try {
                using (var input = new MemoryStream(replacement,
                           writable: false)) {
                    detached.FeedData(input);
                }
                foreach (DataPartReferenceRelationship relationship in
                         relationships) {
                    string targetId = GetNextMediaRelationshipId(owner);
                    AddMediaRelationship(detached, relationship, targetId);
                    replacements.Add((relationship, targetId));
                }
                foreach ((DataPartReferenceRelationship source,
                             string targetId) in replacements) {
                    ReplaceRelationshipReferences(Element, source.Id,
                        targetId);
                }
                foreach (DataPartReferenceRelationship relationship in
                         relationships) {
                    if (CountRelationshipReferences(owner.RootElement,
                            relationship.Id) == 0) {
                        owner.DeleteReferenceRelationship(relationship);
                        deletedOriginals.Add(relationship);
                    }
                }
            } catch {
                foreach ((DataPartReferenceRelationship source,
                             string targetId) in replacements) {
                    ReplaceRelationshipReferences(Element, targetId,
                        source.Id);
                }
                foreach (DataPartReferenceRelationship relationship in
                         owner.DataPartReferenceRelationships.Where(
                             relationship => ReferenceEquals(
                                 relationship.DataPart, detached)).ToArray()) {
                    owner.DeleteReferenceRelationship(relationship);
                }
                foreach (DataPartReferenceRelationship relationship in
                         deletedOriginals) {
                    if (!owner.DataPartReferenceRelationships.Any(
                            candidate => string.Equals(candidate.Id,
                                relationship.Id,
                                StringComparison.Ordinal))) {
                        AddMediaRelationship(original, relationship,
                            relationship.Id);
                    }
                }
                if (!detached.GetDataPartReferenceRelationships().Any()) {
                    document.DeletePart(detached);
                }
                throw;
            }
        }

        private void AddMediaRelationship(MediaDataPart mediaPart,
            DataPartReferenceRelationship source, string relationshipId) {
            SlidePart owner = SlidePart;
            if (source is AudioReferenceRelationship) {
                owner.AddAudioReferenceRelationship(mediaPart,
                    relationshipId);
            } else if (source is VideoReferenceRelationship) {
                owner.AddVideoReferenceRelationship(mediaPart,
                    relationshipId);
            } else {
                owner.AddMediaReferenceRelationship(mediaPart,
                    relationshipId);
            }
        }

        private static void ReplaceRelationshipReferences(
            OpenXmlElement root, string oldValue, string newValue) {
            foreach (OpenXmlElement element in new[] { root }
                         .Concat(root.Descendants())) {
                foreach (OpenXmlAttribute attribute in element
                             .GetAttributes().Where(attribute =>
                                 string.Equals(attribute.Value, oldValue,
                                     StringComparison.Ordinal)).ToArray()) {
                    element.SetAttribute(new OpenXmlAttribute(
                        attribute.Prefix, attribute.LocalName,
                        attribute.NamespaceUri, newValue));
                }
            }
        }

        private static string GetNextMediaRelationshipId(
            SlidePart owner) {
            var used = new HashSet<string>(owner.Parts.Select(pair =>
                    pair.RelationshipId)
                .Concat(owner.ExternalRelationships.Select(item => item.Id))
                .Concat(owner.HyperlinkRelationships.Select(item => item.Id))
                .Concat(owner.DataPartReferenceRelationships.Select(item =>
                    item.Id)), StringComparer.Ordinal);
            int next = 1;
            string relationshipId;
            do {
                relationshipId = "rId" + next++;
            } while (!used.Add(relationshipId));
            return relationshipId;
        }

        internal static bool TryGetMediaKind(Picture picture, out PowerPointMediaKind kind) {
            ApplicationNonVisualDrawingProperties? properties =
                picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;

            if (properties?.GetFirstChild<A.AudioFromFile>() != null) {
                kind = PowerPointMediaKind.Audio;
                return true;
            }

            if (properties?.GetFirstChild<A.VideoFromFile>() != null) {
                kind = PowerPointMediaKind.Video;
                return true;
            }

            kind = default;
            return false;
        }

        private MediaDataPart? GetMediaDataPart() {
            string? relationshipId = MediaReferenceId;
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                return null;
            }

            return SlidePart.DataPartReferenceRelationships
                .FirstOrDefault(rel => rel.Id == relationshipId)?
                .DataPart as MediaDataPart;
        }
    }
}
