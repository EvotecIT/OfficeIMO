using System.Collections.ObjectModel;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordAnimationInfoAtomForWrite = 0x0FF1;
        private const ushort RecordAnimationInfoForWrite = 0x1014;

        internal static bool TryReadClassicAnimations(
            IEnumerable<PowerPointSlide> slides,
            LegacyPptWriterSoundCatalog soundCatalog,
            out LegacyPptWriterAnimationCatalog catalog, out string? reason) {
            if (slides == null) throw new ArgumentNullException(nameof(slides));
            if (soundCatalog == null) throw new ArgumentNullException(nameof(soundCatalog));
            var items = new Dictionary<string, LegacyPptWriterAnimation>(
                StringComparer.Ordinal);
            reason = null;
            foreach (PowerPointSlide slide in slides) {
                IReadOnlyList<PowerPointClassicAnimation> animations =
                    slide.ClassicAnimations;
                if (slide.SlidePart.Slide?.Timing != null
                    && !slide.HasOnlyClassicAnimationTiming()
                    && !HasOnlyDefaultMediaPlaybackTiming(slide)) {
                    catalog = new LegacyPptWriterAnimationCatalog(items);
                    reason = "The slide contains an advanced timing tree that is neither a constrained classic animation sequence nor default embedded-media playback timing.";
                    return false;
                }
                var shapeIds = new HashSet<uint>(slide.EnumerateShapesDeep(
                        slide.Shapes, includeHidden: true)
                    .Select(shape => shape.Id)
                    .Where(id => id.HasValue)
                    .Select(id => id!.Value));
                foreach (PowerPointClassicAnimation animation in animations) {
                    if (!shapeIds.Contains(animation.ShapeId)) {
                        catalog = new LegacyPptWriterAnimationCatalog(items);
                        reason = $"Classic animation target shape {animation.ShapeId} does not exist on its slide.";
                        return false;
                    }
                    if (animation.Order < short.MinValue
                        || animation.Order > short.MaxValue
                        || animation.DelayMilliseconds < 0) {
                        catalog = new LegacyPptWriterAnimationCatalog(items);
                        reason = "A classic animation order or delay lies outside the binary PowerPoint range.";
                        return false;
                    }
                    LegacyPptWriterSound? sound = null;
                    if (animation.PlaysSound) {
                        if (string.IsNullOrWhiteSpace(animation.SoundRelationshipId)) {
                            catalog = new LegacyPptWriterAnimationCatalog(items);
                            reason = "A classic animation is marked to play sound but has no embedded audio relationship.";
                            return false;
                        }
                        var soundElement = new Sound {
                            Embed = animation.SoundRelationshipId,
                            Name = animation.SoundName ?? "Animation Sound"
                        };
                        if (!soundCatalog.TryGetOrAdd(slide.SlidePart,
                                soundElement, out sound, out reason)
                            || sound == null) {
                            catalog = new LegacyPptWriterAnimationCatalog(items);
                            return false;
                        }
                    }
                    string key = GetAnimationKey(slide, animation.ShapeId);
                    if (items.ContainsKey(key)) {
                        catalog = new LegacyPptWriterAnimationCatalog(items);
                        reason = $"Shape {animation.ShapeId} has more than one classic animation.";
                        return false;
                    }
                    items.Add(key, new LegacyPptWriterAnimation(animation, sound));
                }
            }
            catalog = new LegacyPptWriterAnimationCatalog(items);
            return true;
        }

        internal static byte[] BuildAnimationInfoRecord(
            LegacyPptWriterAnimation animation) {
            if (animation == null) throw new ArgumentNullException(nameof(animation));
            var payload = new byte[28];
            WriteUInt32(payload, 0, animation.RawDimColor);
            uint flags = 0;
            if (animation.Reverse) flags |= 1U;
            if (animation.Automatic) flags |= 1U << 2;
            if (animation.Sound != null) flags |= 1U << 4;
            if (animation.StopsSound) flags |= 1U << 6;
            if (animation.AnimateBackground) flags |= 1U << 14;
            WriteUInt32(payload, 4, flags);
            WriteUInt32(payload, 8, animation.Sound?.Id ?? 0U);
            WriteInt32(payload, 12, animation.Automatic
                ? animation.DelayMilliseconds : 0);
            WriteInt16(payload, 16, checked((short)animation.Order));
            WriteUInt16(payload, 18, ushort.MaxValue);
            payload[20] = (byte)animation.BuildType;
            payload[21] = (byte)animation.Effect;
            payload[22] = animation.Direction;
            payload[23] = (byte)animation.AfterEffect;
            payload[24] = (byte)animation.TextBuild;
            byte[] atom = BuildRecord(version: 1, instance: 0,
                RecordAnimationInfoAtomForWrite, payload);
            return BuildContainer(RecordAnimationInfoForWrite, instance: 0,
                new[] { atom });
        }

        private static string GetAnimationKey(PowerPointSlide slide,
            uint shapeId) => slide.SlidePart.Uri + "#" + shapeId.ToString(
                System.Globalization.CultureInfo.InvariantCulture);

        internal sealed class LegacyPptWriterAnimationCatalog {
            private readonly IReadOnlyDictionary<string, LegacyPptWriterAnimation>
                _items;

            internal LegacyPptWriterAnimationCatalog(
                IReadOnlyDictionary<string, LegacyPptWriterAnimation> items) {
                _items = new ReadOnlyDictionary<string, LegacyPptWriterAnimation>(
                    items.ToDictionary(pair => pair.Key, pair => pair.Value,
                        StringComparer.Ordinal));
            }

            internal LegacyPptWriterAnimation? Get(PowerPointShape shape) {
                uint? shapeId = shape.Id;
                PowerPointSlide? slide = shape.OwnerSlide;
                if (!shapeId.HasValue || slide == null) return null;
                return _items.TryGetValue(GetAnimationKey(slide,
                    shapeId.Value), out LegacyPptWriterAnimation? animation)
                    ? animation : null;
            }
        }

        internal sealed class LegacyPptWriterAnimation {
            internal LegacyPptWriterAnimation(
                PowerPointClassicAnimation animation,
                LegacyPptWriterSound? sound) {
                Effect = animation.Effect;
                Direction = animation.Direction;
                BuildType = animation.BuildType;
                Automatic = animation.Automatic;
                DelayMilliseconds = animation.DelayMilliseconds;
                Order = animation.Order;
                Reverse = animation.Reverse;
                AnimateBackground = animation.AnimateBackground;
                AfterEffect = animation.AfterEffect;
                TextBuild = animation.TextBuild;
                RawDimColor = animation.RawDimColor;
                StopsSound = animation.StopsSound;
                Sound = sound;
            }

            internal PowerPointClassicAnimationEffect Effect { get; }
            internal byte Direction { get; }
            internal PowerPointClassicAnimationBuildType BuildType { get; }
            internal bool Automatic { get; }
            internal int DelayMilliseconds { get; }
            internal int Order { get; }
            internal bool Reverse { get; }
            internal bool AnimateBackground { get; }
            internal PowerPointClassicAnimationAfterEffect AfterEffect { get; }
            internal PowerPointClassicTextBuild TextBuild { get; }
            internal uint RawDimColor { get; }
            internal bool StopsSound { get; }
            internal LegacyPptWriterSound? Sound { get; }
        }
    }
}
