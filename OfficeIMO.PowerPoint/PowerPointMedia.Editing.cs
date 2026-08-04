using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    /// <summary>Storage kind of a PowerPoint media frame.</summary>
    public enum PowerPointMediaSourceKind {
        /// <summary>Payload is stored in the package.</summary>
        Embedded,
        /// <summary>Payload is referenced by an external URI.</summary>
        Linked,
        /// <summary>The relationship cannot be resolved.</summary>
        Broken
    }

    /// <summary>Typed playback metadata shared by audio and video.</summary>
    public sealed class PowerPointMediaPlaybackOptions {
        /// <summary>Volume from zero through 100 percent.</summary>
        public int VolumePercent { get; set; } = 80;
        /// <summary>Whether audio is muted.</summary>
        public bool Mute { get; set; }
        /// <summary>Whether playback repeats indefinitely.</summary>
        public bool Loop { get; set; }
        /// <summary>Whether the media frame remains visible when stopped.</summary>
        public bool ShowWhenStopped { get; set; } = true;
        /// <summary>Whether video plays full-screen.</summary>
        public bool FullScreen { get; set; }
        /// <summary>Trim start in milliseconds.</summary>
        public uint? TrimStartMilliseconds { get; set; }
        /// <summary>Trim end in milliseconds.</summary>
        public uint? TrimEndMilliseconds { get; set; }
        /// <summary>Fade-in duration in milliseconds.</summary>
        public uint? FadeInMilliseconds { get; set; }
        /// <summary>Fade-out duration in milliseconds.</summary>
        public uint? FadeOutMilliseconds { get; set; }
    }

    public partial class PowerPointMedia {
        /// <summary>Gets whether the payload is embedded, linked, or broken.</summary>
        public PowerPointMediaSourceKind SourceKind {
            get {
                string? id = MediaReferenceId;
                if (string.IsNullOrWhiteSpace(id)) return PowerPointMediaSourceKind.Broken;
                if (GetMediaDataPart() != null) return PowerPointMediaSourceKind.Embedded;
                return SlidePart.ExternalRelationships.Any(item => item.Id == id)
                    ? PowerPointMediaSourceKind.Linked : PowerPointMediaSourceKind.Broken;
            }
        }

        /// <summary>Gets the external media URI when linked.</summary>
        public Uri? LinkUri {
            get {
                string? id = MediaReferenceId;
                return SlidePart.ExternalRelationships.FirstOrDefault(item => item.Id == id)?.Uri;
            }
        }

        /// <summary>Updates a linked media URI without rewriting the frame or timing tree.</summary>
        public PowerPointMedia UpdateLink(Uri uri) {
            if (uri == null) throw new ArgumentNullException(nameof(uri));
            if (!uri.IsAbsoluteUri) throw new ArgumentException(
                "A linked media URI must be absolute.", nameof(uri));
            string id = MediaReferenceId ?? throw new InvalidOperationException("The media frame has no file relationship.");
            ExternalRelationship relationship = SlidePart.ExternalRelationships
                .FirstOrDefault(item => item.Id == id)
                ?? throw new InvalidOperationException("The media frame is embedded rather than linked.");
            SlidePart.DeleteExternalRelationship(id);
            SlidePart.AddExternalRelationship(relationship.RelationshipType, uri, id);
            return this;
        }

        /// <summary>Reads typed playback metadata from the native timing and media nodes.</summary>
        public PowerPointMediaPlaybackOptions GetPlaybackOptions() {
            CommonMediaNode? node = FindCommonMediaNode();
            P14.Media? metadata = ((Picture)Element).Descendants<P14.Media>().FirstOrDefault();
            P14.MediaTrim? trim = metadata?.GetFirstChild<P14.MediaTrim>();
            P14.MediaFade? fade = metadata?.GetFirstChild<P14.MediaFade>();
            Video? video = node?.Parent as Video;
            return new PowerPointMediaPlaybackOptions {
                VolumePercent = (int)Math.Round((node?.Volume?.Value ?? 80000) / 1000D),
                Mute = node?.Mute?.Value == true,
                Loop = string.Equals(node?.CommonTimeNode?.RepeatCount?.Value,
                    "indefinite", StringComparison.OrdinalIgnoreCase),
                ShowWhenStopped = node?.ShowWhenStopped?.Value != false,
                FullScreen = video?.FullScreen?.Value == true,
                TrimStartMilliseconds = ParseMilliseconds(trim?.Start?.Value),
                TrimEndMilliseconds = ParseMilliseconds(trim?.End?.Value),
                FadeInMilliseconds = ParseMilliseconds(fade?.InDuration?.Value),
                FadeOutMilliseconds = ParseMilliseconds(fade?.OutDuration?.Value)
            };
        }

        /// <summary>Updates native playback, trim, and fade metadata in place.</summary>
        public PowerPointMedia SetPlaybackOptions(PowerPointMediaPlaybackOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (options.VolumePercent < 0 || options.VolumePercent > 100)
                throw new ArgumentOutOfRangeException(nameof(options), "Volume must be between zero and 100 percent.");
            if (options.TrimStartMilliseconds.HasValue && options.TrimEndMilliseconds.HasValue &&
                options.TrimStartMilliseconds.Value >= options.TrimEndMilliseconds.Value)
                throw new ArgumentException("Trim start must be earlier than trim end.", nameof(options));
            CommonMediaNode node = FindCommonMediaNode()
                ?? throw new InvalidOperationException("The media frame has no editable playback timing node.");
            P14.Media metadata = ((Picture)Element).Descendants<P14.Media>().FirstOrDefault()
                ?? throw new InvalidOperationException(
                    "The producer-specific media frame has no editable p14 playback metadata and was left unchanged.");
            node.Volume = options.VolumePercent * 1000;
            node.Mute = options.Mute;
            node.ShowWhenStopped = options.ShowWhenStopped;
            node.CommonTimeNode!.RepeatCount = options.Loop ? "indefinite" : null;
            if (node.Parent is Video video) video.FullScreen = options.FullScreen;

            metadata.GetFirstChild<P14.MediaTrim>()?.Remove();
            if (options.TrimStartMilliseconds.HasValue || options.TrimEndMilliseconds.HasValue) {
                metadata.PrependChild(new P14.MediaTrim {
                    Start = options.TrimStartMilliseconds?.ToString(CultureInfo.InvariantCulture),
                    End = options.TrimEndMilliseconds?.ToString(CultureInfo.InvariantCulture)
                });
            }
            metadata.GetFirstChild<P14.MediaFade>()?.Remove();
            if (options.FadeInMilliseconds.HasValue || options.FadeOutMilliseconds.HasValue) {
                metadata.Append(new P14.MediaFade {
                    InDuration = options.FadeInMilliseconds?.ToString(CultureInfo.InvariantCulture),
                    OutDuration = options.FadeOutMilliseconds?.ToString(CultureInfo.InvariantCulture)
                });
            }
            return this;
        }

        private CommonMediaNode? FindCommonMediaNode() {
            string shapeId = (Id ?? 0U).ToString(CultureInfo.InvariantCulture);
            return SlidePart.Slide?.Timing?.Descendants<CommonMediaNode>()
                .FirstOrDefault(node => node.GetFirstChild<TargetElement>()?
                    .GetFirstChild<ShapeTarget>()?.ShapeId?.Value == shapeId);
        }

        private static uint? ParseMilliseconds(string? value) =>
            uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture,
                out uint parsed) ? parsed : null;
    }
}
