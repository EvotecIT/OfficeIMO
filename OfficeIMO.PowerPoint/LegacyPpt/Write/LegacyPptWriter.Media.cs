using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordExternalMediaAtomForWrite = 0x1004;
        private const ushort RecordExternalWavAudioEmbeddedForWrite = 0x100F;
        private const ushort RecordExternalWavAudioEmbeddedAtomForWrite = 0x1013;

        internal static bool TryReadMedia(IEnumerable<PowerPointSlide> slides,
            uint firstObjectId, LegacyPptWriterSoundCatalog soundCatalog,
            out LegacyPptWriterMediaCatalog catalog, out string? reason) {
            if (slides == null) throw new ArgumentNullException(nameof(slides));
            if (soundCatalog == null) {
                throw new ArgumentNullException(nameof(soundCatalog));
            }
            catalog = new LegacyPptWriterMediaCatalog(firstObjectId);
            reason = null;
            foreach (PowerPointSlide slide in slides) {
                PowerPointMedia[] mediaShapes = slide.EnumerateShapesDeep(
                        slide.Shapes, includeHidden: true)
                    .OfType<PowerPointMedia>().ToArray();
                if (mediaShapes.Length > 0
                    && !HasOnlyDefaultMediaPlaybackTiming(slide)) {
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
                    if (!soundCatalog.TryGetOrAddMedia(slide.SlidePart,
                            media, out LegacyPptWriterSound? sound,
                            out reason) || sound == null) {
                        catalog = new LegacyPptWriterMediaCatalog(firstObjectId);
                        return false;
                    }
                    catalog.Add(media.Element, new LegacyPptWriterMedia(
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
            PowerPointSlide slide) {
            Timing? timing = slide.SlidePart.Slide?.Timing;
            if (timing == null) return false;
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
            ShapeTarget[] targets = timing.Descendants<ShapeTarget>()
                .ToArray();
            CommonMediaNode[] mediaNodes = timing
                .Descendants<CommonMediaNode>().ToArray();
            return targets.Length == mediaShapeIds.Count
                && mediaNodes.Length == mediaShapeIds.Count
                && timing.Descendants().All(element =>
                    allowedNames.Contains(element.LocalName))
                && targets.All(target =>
                    !string.IsNullOrWhiteSpace(target.ShapeId?.Value)
                    && mediaShapeIds.Contains(target.ShapeId!.Value!)
                    && target.Ancestors<CommonMediaNode>().Any()
                    && target.Ancestors<Audio>().Any())
                && mediaNodes.All(HasDefaultMediaNode);
        }

        private static bool HasDefaultMediaNode(CommonMediaNode mediaNode) {
            OpenXmlAttribute[] mediaAttributes = mediaNode.GetAttributes()
                .ToArray();
            if (mediaAttributes.Length != 1
                || !string.Equals(mediaAttributes[0].LocalName, "vol",
                    StringComparison.Ordinal)
                || !string.Equals(mediaAttributes[0].Value, "80000",
                    StringComparison.Ordinal)) {
                return false;
            }
            CommonTimeNode? commonTime = mediaNode
                .GetFirstChild<CommonTimeNode>();
            TargetElement? target = mediaNode.GetFirstChild<TargetElement>();
            if (commonTime == null || target == null
                || mediaNode.ChildElements.Count != 2) return false;
            IReadOnlyDictionary<string, string?> timeAttributes = commonTime
                .GetAttributes().ToDictionary(attribute => attribute.LocalName,
                    attribute => attribute.Value, StringComparer.Ordinal);
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
            StartConditionList? conditions = commonTime
                .GetFirstChild<StartConditionList>();
            Condition? condition = conditions?.GetFirstChild<Condition>();
            return commonTime.ChildElements.Count == 1
                && conditions?.ChildElements.Count == 1
                && condition != null
                && condition.GetAttributes().Count == 1
                && string.Equals(condition.GetAttributes()[0].LocalName,
                    "delay", StringComparison.Ordinal)
                && string.Equals(condition.GetAttributes()[0].Value,
                    "indefinite", StringComparison.Ordinal);
        }

        internal sealed class LegacyPptWriterMediaCatalog {
            private readonly Dictionary<OpenXmlElement,
                LegacyPptWriterMedia> _mediaByShape = new(
                    ReferenceComparer.Instance);
            private readonly List<LegacyPptWriterMedia> _media = new();
            private uint _nextObjectId;

            internal LegacyPptWriterMediaCatalog(uint firstObjectId = 1) {
                _nextObjectId = firstObjectId;
            }

            internal IReadOnlyList<LegacyPptWriterMedia> Media =>
                new ReadOnlyCollection<LegacyPptWriterMedia>(_media);

            internal uint NextObjectId => _nextObjectId;

            internal LegacyPptWriterMedia? Get(PowerPointShape shape) =>
                _mediaByShape.TryGetValue(shape.Element,
                    out LegacyPptWriterMedia? value) ? value : null;

            internal void Add(OpenXmlElement shape,
                LegacyPptWriterMedia media) {
                _mediaByShape.Add(shape, media);
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
