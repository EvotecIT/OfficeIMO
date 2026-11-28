using System.Globalization;
using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

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
                if (TryGetPackageElement(out _, out _)) {
                    if (!GetPackageMediaPart().IsNone) return PowerPointMediaSourceKind.Embedded;
                    return TryGetPackageMediaLink(id!, out _)
                        ? PowerPointMediaSourceKind.Linked : PowerPointMediaSourceKind.Broken;
                }
                return GetSdkMediaSourceKind(id!);
            }
        }

        /// <summary>Gets the external media URI when linked.</summary>
        public Uri? LinkUri {
            get {
                string? id = MediaReferenceId;
                if (TryGetPackageElement(out _, out _))
                    return id != null && TryGetPackageMediaLink(id, out PowerPointRelationship link)
                        && Uri.TryCreate(link.Target, UriKind.RelativeOrAbsolute, out Uri? uri) ? uri : null;
                return GetSdkLinkedMediaUri(id);
            }
        }

        private bool TryGetPackageMediaLink(string id, out PowerPointRelationship relationship) =>
            SlidePartRef.TryGetRelationship(id, out relationship) && relationship.IsExternal
            && relationship.Type != PowerPointPartRef.HyperlinkRelationshipType;

        /// <summary>Updates a linked media URI without rewriting the frame or timing tree.</summary>
        public PowerPointMedia UpdateLink(Uri uri) {
            ArgumentNullException.ThrowIfNull(uri);
            if (!uri.IsAbsoluteUri) throw new ArgumentException(
                "A linked media URI must be absolute.", nameof(uri));
            string id = MediaReferenceId ?? throw new InvalidOperationException("The media frame has no file relationship.");
            PowerPointPartRef part = SlidePartRef;
            // The SDK retains the supplied absolute Uri in memory when a file path omits the file: scheme.
            bool implicitFileUri = uri.IsFile && !uri.OriginalString.StartsWith("file:", StringComparison.OrdinalIgnoreCase);
            if (!implicitFileUri && !part.IsSdk && part.Document is { EditablePackage: { } package } document) {
                if (!TryGetPackageMediaLink(id, out _)) throw new InvalidOperationException("The media frame is embedded rather than linked.");
                OfficeXmlTree relationships = package.GetRelationships(part.Name);
                OfficeXmlNode root = new(relationships, relationships.Root);
                foreach (OfficeXmlNode relationshipNode in root.Children(OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                    if (relationshipNode.Attribute("Id") != id) continue;
                    relationshipNode.SetAttribute("Target", PowerPointHyperlinkResolver.TargetText(uri));
                    part.MarkEdited();
                    document.InvalidatePackageRelationships();
                    return this;
                }
                throw new InvalidOperationException("The media frame has no file relationship.");
            }
            UpdateSdkMediaLink(id, uri);
            return this;
        }

        /// <summary>Reads typed playback metadata from the native timing and media nodes.</summary>
        public PowerPointMediaPlaybackOptions GetPlaybackOptions() {
            const int p = OfficeXmlNamespaces.PresentationML;
            OfficeXmlNode node = FindCommonMediaNode(SlidePartRef);
            OfficeXmlNode metadata = PlaybackMetadata;
            OfficeXmlNode trim = metadata.Child(PlaybackMediaNamespace, "trim");
            OfficeXmlNode fade = metadata.Child(PlaybackMediaNamespace, "fade");
            OfficeXmlNode video = node.Parent;
            int volume = node.Attribute("vol") is { } value ? XmlConvert.ToInt32(value) : 80000;
            return new PowerPointMediaPlaybackOptions {
                VolumePercent = (int)Math.Round(volume / 1000D),
                Mute = ReadPlaybackBoolean(node, "mute") == true,
                Loop = string.Equals(node.Child(p, "cTn").Attribute("repeatCount"), "indefinite", StringComparison.OrdinalIgnoreCase),
                ShowWhenStopped = ReadPlaybackBoolean(node, "showWhenStopped") != false,
                FullScreen = video.Is(p, "video") && ReadPlaybackBoolean(video, "fullScrn") == true,
                TrimStartMilliseconds = ParseMilliseconds(trim.Attribute("st")),
                TrimEndMilliseconds = ParseMilliseconds(trim.Attribute("end")),
                FadeInMilliseconds = ParseMilliseconds(fade.Attribute("in")),
                FadeOutMilliseconds = ParseMilliseconds(fade.Attribute("out"))
            };
        }

        /// <summary>Updates native playback, trim, and fade metadata in place.</summary>
        public PowerPointMedia SetPlaybackOptions(PowerPointMediaPlaybackOptions options) {
            ArgumentNullException.ThrowIfNull(options);
            if (options.VolumePercent < 0 || options.VolumePercent > 100)
                throw new ArgumentOutOfRangeException(nameof(options), "Volume must be between zero and 100 percent.");
            if (options.TrimStartMilliseconds.HasValue && options.TrimEndMilliseconds.HasValue &&
                options.TrimStartMilliseconds.Value >= options.TrimEndMilliseconds.Value)
                throw new ArgumentException("Trim start must be earlier than trim end.", nameof(options));
            const int p = OfficeXmlNamespaces.PresentationML;
            PowerPointPartRef part = SlidePartRef.ForMovingXmlNodes();
            OfficeXmlNode node = FindCommonMediaNode(part);
            OfficeXmlNode timing = node.Child(p, "cTn");
            if (timing.IsNone) throw new InvalidOperationException("The media frame has no editable playback timing node.");
            OfficeXmlNode metadata = PlaybackMetadata;
            if (metadata.IsNone) throw new InvalidOperationException(
                "The producer-specific media frame has no editable p14 playback metadata and was left unchanged.");

            node.SetInt32Attribute("vol", options.VolumePercent * 1000);
            node.SetBooleanAttribute("mute", options.Mute);
            node.SetBooleanAttribute("showWhenStopped", options.ShowWhenStopped);
            timing.SetAttribute("repeatCount", options.Loop ? "indefinite" : null);
            OfficeXmlNode video = node.Parent;
            if (video.Is(p, "video")) video.SetBooleanAttribute("fullScrn", options.FullScreen);

            metadata.Child(PlaybackMediaNamespace, "trim").Remove();
            if (options.TrimStartMilliseconds.HasValue || options.TrimEndMilliseconds.HasValue) {
                OfficeXmlNode trim = metadata.CreateElement(PlaybackMediaNamespace, "trim", "p14");
                trim.SetUInt32Attribute("st", options.TrimStartMilliseconds);
                trim.SetUInt32Attribute("end", options.TrimEndMilliseconds);
                metadata.InsertBefore(trim, metadata.FirstChild);
            }
            metadata.Child(PlaybackMediaNamespace, "fade").Remove();
            if (options.FadeInMilliseconds.HasValue || options.FadeOutMilliseconds.HasValue) {
                // The SDK appends a replaced fade after any existing bookmarks and extension list.
                OfficeXmlNode fade = metadata.CreateElement(PlaybackMediaNamespace, "fade", "p14");
                fade.SetUInt32Attribute("in", options.FadeInMilliseconds);
                fade.SetUInt32Attribute("out", options.FadeOutMilliseconds);
                metadata.Append(fade);
            }
            part.MarkEdited();
            return this;
        }

        private OfficeXmlNode FindCommonMediaNode(PowerPointPartRef part) =>
            PowerPointSlide.FindMediaTimingNode(part.Root, (Id ?? 0U).ToString(CultureInfo.InvariantCulture));

        private static bool? ReadPlaybackBoolean(OfficeXmlNode node, string name) =>
            node.Attribute(name) is { } value ? XmlConvert.ToBoolean(value) : null;

        private static uint? ParseMilliseconds(string? value) =>
            uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture,
                out uint parsed) ? parsed : null;
    }
}
