using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
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
