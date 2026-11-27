using System.Collections.ObjectModel;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using static OfficeIMO.PowerPoint.LegacyPpt.Write.LegacyPptWriterXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordExternalMediaAtomForWrite = 0x1004;
        private const ushort RecordExternalWavAudioEmbeddedForWrite = 0x100F;
        private const ushort RecordExternalWavAudioEmbeddedAtomForWrite = 0x1013;

        internal static bool TryReadMedia(LegacyPptPackageView view,
            IEnumerable<PowerPointSlide> slides,
            uint firstObjectId, LegacyPptWriterSoundCatalog soundCatalog,
            out LegacyPptWriterMediaCatalog catalog, out string? reason) {
            ArgumentNullException.ThrowIfNull(view);
            ArgumentNullException.ThrowIfNull(slides);
            ArgumentNullException.ThrowIfNull(soundCatalog);
            catalog = new LegacyPptWriterMediaCatalog(firstObjectId);
            reason = null;
            foreach (PowerPointSlide slide in slides) {
                LegacyPptPartView slidePart = view.Part(slide.PartRef.Name);
                PowerPointMedia[] mediaShapes = slide.EnumerateShapesDeep(
                        slide.Shapes, includeHidden: true)
                    .OfType<PowerPointMedia>().ToArray();
                if (mediaShapes.Length > 0
                    && !HasOnlyDefaultMediaPlaybackTiming(slidePart, slide)) {
                    reason = "Binary embedded media requires the default media-only playback timing tree.";
                    catalog = new LegacyPptWriterMediaCatalog(firstObjectId);
                    return false;
                }
                foreach (PowerPointMedia media in mediaShapes) {
                    if (media.Kind != PowerPointMediaKind.Audio) {
                        reason = "PowerPoint 97-2003 has no native embedded-video representation; video conversion is blocked.";
                        catalog = new LegacyPptWriterMediaCatalog(firstObjectId);
                        return false;
                    }
                    if (!soundCatalog.TryGetOrAddMedia(slidePart,
                            media, out LegacyPptWriterSound? sound,
                            out reason) || sound == null) {
                        catalog = new LegacyPptWriterMediaCatalog(firstObjectId);
                        return false;
                    }
                    catalog.Add(media, new LegacyPptWriterMedia(
                        catalog.NextObjectId, sound));
                }
            }
            return true;
        }

        internal static byte[] BuildExternalMediaRecord(
            LegacyPptWriterMedia media) {
            var mediaPayload = new byte[8];
            WriteUInt32(mediaPayload, 0, media.Id);
            var wavePayload = new byte[8];
            WriteUInt32(wavePayload, 0, media.Sound.Id);
            WriteInt32(wavePayload, 4, media.DurationMilliseconds);
            return BuildContainer(RecordExternalWavAudioEmbeddedForWrite,
                instance: 0, new[] {
                    BuildRecord(version: 0, instance: 0,
                        RecordExternalMediaAtomForWrite, mediaPayload),
                    BuildRecord(version: 1, instance: 1,
                        RecordExternalWavAudioEmbeddedAtomForWrite,
                        wavePayload)
                });
        }

        private static bool HasOnlyDefaultMediaPlaybackTiming(
            LegacyPptPartView slidePart, PowerPointSlide slide) {
            OfficeXmlNode timing = slidePart.Root.Child(LegacyPptXml.P,
                "timing");
            if (timing.IsNone) return false;
            var mediaShapeIds = new HashSet<string>(slide
                .EnumerateShapesDeep(slide.Shapes, includeHidden: true)
                .OfType<PowerPointMedia>()
                .Select(item => item.Id?.ToString(
                    System.Globalization.CultureInfo.InvariantCulture))
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Select(id => id!), StringComparer.Ordinal);
            if (mediaShapeIds.Count == 0) return false;
            var allowedNames = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase) {
                "audio", "cMediaNode", "cTn", "childTnLst", "cond",
                "par", "spTgt", "stCondLst", "tgtEl", "tnLst"
            };
            List<OfficeXmlNode> targets = timing.Descendants(LegacyPptXml.P,
                "spTgt");
            List<OfficeXmlNode> mediaNodes = timing.Descendants(LegacyPptXml.P,
                "cMediaNode");
            return targets.Count == mediaShapeIds.Count
                && mediaNodes.Count == mediaShapeIds.Count
                && timing.AllDescendants().All(element =>
                    allowedNames.Contains(element.LocalName))
                && targets.All(target =>
                    target.Attribute("spid") is string shapeId
                    && !string.IsNullOrWhiteSpace(shapeId)
                    && mediaShapeIds.Contains(shapeId)
                    && target.HasAncestor(LegacyPptXml.P, "cMediaNode")
                    && target.HasAncestor(LegacyPptXml.P, "audio"))
                && mediaNodes.All(HasDefaultMediaNode);
        }

        private static bool HasDefaultMediaNode(OfficeXmlNode mediaNode) {
            List<(int NamespaceId, string LocalName, string Value)>
                mediaAttributes = Attributes(mediaNode);
            if (mediaAttributes.Count != 1
                || !string.Equals(mediaAttributes[0].LocalName, "vol",
                    StringComparison.Ordinal)
                || !string.Equals(mediaAttributes[0].Value, "80000",
                    StringComparison.Ordinal)) {
                return false;
            }
            OfficeXmlNode commonTime = mediaNode.Child(LegacyPptXml.P, "cTn");
            OfficeXmlNode target = mediaNode.Child(LegacyPptXml.P, "tgtEl");
            if (commonTime.IsNone || target.IsNone
                || ChildCount(mediaNode) != 2) return false;
            var timeAttributes = new Dictionary<string, string?>(
                StringComparer.Ordinal);
            foreach ((_, string localName, string value) in
                     Attributes(commonTime)) {
                timeAttributes[localName] = value;
            }
            if (timeAttributes.Count != 3
                || !timeAttributes.TryGetValue("id", out string? id)
                || !uint.TryParse(id,
                    System.Globalization.NumberStyles.None,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out uint parsedId) || parsedId == 0
                || !timeAttributes.TryGetValue("fill", out string? fill)
                || !string.Equals(fill, "hold", StringComparison.Ordinal)
                || !timeAttributes.TryGetValue("display", out string? display)
                || display != "0" && !string.Equals(display, "false",
                    StringComparison.OrdinalIgnoreCase)) {
                return false;
            }
            OfficeXmlNode conditions = commonTime.Child(LegacyPptXml.P,
                "stCondLst");
            OfficeXmlNode condition = conditions.Child(LegacyPptXml.P, "cond");
            List<(int NamespaceId, string LocalName, string Value)>
                conditionAttributes = Attributes(condition);
            return ChildCount(commonTime) == 1
                && !conditions.IsNone && ChildCount(conditions) == 1
                && !condition.IsNone
                && conditionAttributes.Count == 1
                && string.Equals(conditionAttributes[0].LocalName,
                    "delay", StringComparison.Ordinal)
                && string.Equals(conditionAttributes[0].Value,
                    "indefinite", StringComparison.Ordinal);
        }

        internal sealed class LegacyPptWriterMediaCatalog {
            // Keyed by the media shape's write key.
            private readonly Dictionary<object,
                LegacyPptWriterMedia> _mediaByShape = new();
            private readonly List<LegacyPptWriterMedia> _media = new();
            private uint _nextObjectId;

            internal LegacyPptWriterMediaCatalog(uint firstObjectId = 1) {
                _nextObjectId = firstObjectId;
            }

            internal IReadOnlyList<LegacyPptWriterMedia> Media =>
                new ReadOnlyCollection<LegacyPptWriterMedia>(_media);

            internal uint NextObjectId => _nextObjectId;

            internal LegacyPptWriterMedia? Get(PowerPointShape shape) =>
                _mediaByShape.TryGetValue(ShapeKey(shape),
                    out LegacyPptWriterMedia? value) ? value : null;

            internal void Add(PowerPointShape shape,
                LegacyPptWriterMedia media) {
                _mediaByShape.Add(ShapeKey(shape), media);
                _media.Add(media);
                _nextObjectId = checked(media.Id + 1U);
            }
        }

        internal sealed class LegacyPptWriterMedia {
            internal LegacyPptWriterMedia(uint id,
                LegacyPptWriterSound sound) {
                if (id == 0) throw new ArgumentOutOfRangeException(nameof(id));
                Id = id;
                Sound = sound ?? throw new ArgumentNullException(nameof(sound));
                DurationMilliseconds = ReadWaveDurationMilliseconds(
                    sound.DataBytes);
            }

            internal uint Id { get; }
            internal LegacyPptWriterSound Sound { get; }
            internal int DurationMilliseconds { get; }

            private static int ReadWaveDurationMilliseconds(byte[] bytes) {
                if (bytes.Length < 20
                    || bytes[0] != (byte)'R' || bytes[1] != (byte)'I'
                    || bytes[2] != (byte)'F' || bytes[3] != (byte)'F'
                    || bytes[8] != (byte)'W' || bytes[9] != (byte)'A'
                    || bytes[10] != (byte)'V' || bytes[11] != (byte)'E') {
                    return 0;
                }
                uint byteRate = 0;
                uint dataLength = 0;
                for (int offset = 12; offset <= bytes.Length - 8;) {
                    uint chunkLength = ReadUInt32(bytes, offset + 4);
                    long next = (long)offset + 8L + chunkLength
                        + (chunkLength & 1U);
                    if (next > bytes.Length) return 0;
                    if (bytes[offset] == (byte)'f'
                        && bytes[offset + 1] == (byte)'m'
                        && bytes[offset + 2] == (byte)'t'
                        && bytes[offset + 3] == (byte)' '
                        && chunkLength >= 12U) {
                        byteRate = ReadUInt32(bytes, offset + 16);
                    } else if (bytes[offset] == (byte)'d'
                               && bytes[offset + 1] == (byte)'a'
                               && bytes[offset + 2] == (byte)'t'
                               && bytes[offset + 3] == (byte)'a') {
                        dataLength = chunkLength;
                    }
                    if (byteRate != 0 && dataLength != 0) break;
                    offset = checked((int)next);
                }
                if (byteRate == 0 || dataLength == 0) return 0;
                double duration = dataLength * 1000D / byteRate;
                return duration >= int.MaxValue
                    ? int.MaxValue
                    : checked((int)Math.Round(duration,
                        MidpointRounding.AwayFromZero));
            }
        }
    }
}
