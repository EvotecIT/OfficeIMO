using Model = OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal static bool TryReadTransition(PowerPointSlide slide,
            out LegacyPptWriterTransition? transition, out string? reason) {
            transition = null;
            reason = null;
            if (slide.Transition == SlideTransition.None) return true;

            byte effectType;
            byte effectDirection;
            switch (slide.Transition) {
                case SlideTransition.Cut:
                    effectType = 0;
                    effectDirection = 1;
                    break;
                case SlideTransition.Fade:
                    effectType = 6;
                    effectDirection = 0;
                    break;
                case SlideTransition.Wipe:
                    effectType = 10;
                    effectDirection = 0;
                    break;
                case SlideTransition.BlindsVertical:
                    effectType = 2;
                    effectDirection = 0;
                    break;
                case SlideTransition.BlindsHorizontal:
                    effectType = 2;
                    effectDirection = 1;
                    break;
                case SlideTransition.CombHorizontal:
                    effectType = 21;
                    effectDirection = 0;
                    break;
                case SlideTransition.CombVertical:
                    effectType = 21;
                    effectDirection = 1;
                    break;
                case SlideTransition.PushLeft:
                    effectType = 20;
                    effectDirection = 0;
                    break;
                case SlideTransition.PushUp:
                    effectType = 20;
                    effectDirection = 1;
                    break;
                case SlideTransition.PushRight:
                    effectType = 20;
                    effectDirection = 2;
                    break;
                case SlideTransition.PushDown:
                    effectType = 20;
                    effectDirection = 3;
                    break;
                default:
                    reason = $"The {slide.Transition} transition has no PowerPoint 97-2003 representation.";
                    return false;
            }

            byte speed = slide.TransitionSpeed switch {
                SlideTransitionSpeed.Slow => 0,
                SlideTransitionSpeed.Fast => 2,
                _ => 1
            };
            if (slide.TransitionDurationSeconds.HasValue) {
                double duration = slide.TransitionDurationSeconds.Value;
                byte? durationSpeed = NearlyEqual(duration, 0.75) ? (byte)0
                    : NearlyEqual(duration, 0.5) ? (byte)1
                    : NearlyEqual(duration, 0.25) ? (byte)2
                    : null;
                if (!durationSpeed.HasValue) {
                    reason = "Binary PowerPoint transition duration must be exactly 0.75, 0.5, or 0.25 seconds.";
                    return false;
                }
                if (slide.TransitionSpeed.HasValue && speed != durationSpeed.Value) {
                    reason = "The requested transition speed and duration describe different binary speed values.";
                    return false;
                }
                speed = durationSpeed.Value;
            }

            int slideTime = 0;
            bool autoAdvance = slide.TransitionAdvanceAfterSeconds.HasValue;
            if (autoAdvance) {
                double seconds = slide.TransitionAdvanceAfterSeconds!.Value;
                if (seconds < 0 || seconds > 86399) {
                    reason = "Binary PowerPoint automatic advance must be between 0 and 86,399 seconds.";
                    return false;
                }
                slideTime = checked((int)Math.Round(seconds * 1000.0,
                    MidpointRounding.AwayFromZero));
            }
            transition = new LegacyPptWriterTransition(effectType, effectDirection,
                speed, slide.TransitionAdvanceOnClick != false, autoAdvance,
                slideTime);
            return true;
        }

        private static bool NearlyEqual(double left, double right) =>
            Math.Abs(left - right) < 0.0005;

        internal static byte[] PatchSlideShowInfo(byte[] record,
            PowerPointSlide slide) {
            if (record.Length < 24) {
                throw new InvalidDataException(
                    "The slide-show information atom is too short.");
            }
            if (!TryReadTransition(slide, out LegacyPptWriterTransition? transition,
                    out string? reason)) {
                throw new NotSupportedException(reason);
            }
            PatchSlideShowInfoPayload(record, slide.Hidden, transition);
            return record;
        }

        internal static byte[] BuildSlideShowInfoRecord(PowerPointSlide slide) {
            if (!TryReadTransition(slide, out LegacyPptWriterTransition? transition,
                    out string? reason)) {
                throw new NotSupportedException(reason);
            }
            var payload = new byte[16];
            PatchSlideShowInfoPayload(payload, slide.Hidden, transition,
                payloadOffset: 0);
            return BuildRecord(version: 0, instance: 0,
                RecordSlideShowSlideInfoAtom, payload);
        }

        private static void PatchSlideShowInfoPayload(byte[] bytes, bool hidden,
            LegacyPptWriterTransition? transition, int payloadOffset = 8) {
            WriteInt32(bytes, payloadOffset, transition?.SlideTimeMilliseconds ?? 0);
            bytes[payloadOffset + 8] = transition?.EffectDirection ?? 0;
            bytes[payloadOffset + 9] = transition?.EffectType ?? 0;
            ushort flags = ReadUInt16(bytes, payloadOffset + 10);
            flags = unchecked((ushort)(flags & ~(0x0001 | 0x0004 | 0x0400)));
            if (hidden) flags |= 0x0004;
            if (transition?.ManualAdvance != false) flags |= 0x0001;
            if (transition?.AutoAdvance == true) flags |= 0x0400;
            WriteUInt16(bytes, payloadOffset + 10, flags);
            bytes[payloadOffset + 12] = transition?.Speed ?? 1;
        }

        internal sealed class LegacyPptWriterTransition {
            internal LegacyPptWriterTransition(byte effectType, byte effectDirection,
                byte speed, bool manualAdvance, bool autoAdvance,
                int slideTimeMilliseconds) {
                EffectType = effectType;
                EffectDirection = effectDirection;
                Speed = speed;
                ManualAdvance = manualAdvance;
                AutoAdvance = autoAdvance;
                SlideTimeMilliseconds = slideTimeMilliseconds;
            }

            internal byte EffectType { get; }
            internal byte EffectDirection { get; }
            internal byte Speed { get; }
            internal bool ManualAdvance { get; }
            internal bool AutoAdvance { get; }
            internal int SlideTimeMilliseconds { get; }

            internal static LegacyPptWriterTransition? FromLegacyProjection(
                OfficeIMO.PowerPoint.LegacyPpt.Model.LegacyPptTransition? source) {
                if (source == null) return null;
                byte effectType;
                byte effectDirection;
                switch (source.Effect) {
                    case Model.LegacyPptTransitionEffect.Cut:
                        effectType = 0;
                        effectDirection = 1;
                        break;
                    case Model.LegacyPptTransitionEffect.Fade:
                        effectType = 6;
                        effectDirection = 0;
                        break;
                    case Model.LegacyPptTransitionEffect.Wipe:
                        effectType = 10;
                        effectDirection = 0;
                        break;
                    case Model.LegacyPptTransitionEffect.Blinds:
                        effectType = 2;
                        effectDirection = source.EffectDirection == 0 ? (byte)0 : (byte)1;
                        break;
                    case Model.LegacyPptTransitionEffect.Comb:
                        effectType = 21;
                        effectDirection = source.EffectDirection == 0 ? (byte)0 : (byte)1;
                        break;
                    case Model.LegacyPptTransitionEffect.Push:
                        effectType = 20;
                        effectDirection = source.EffectDirection <= 3
                            ? source.EffectDirection
                            : (byte)0;
                        break;
                    default:
                        return null;
                }
                byte speed = source.Speed <= 2 ? source.Speed : (byte)1;
                return new LegacyPptWriterTransition(effectType, effectDirection,
                    speed, source.ManualAdvance, source.AutoAdvance,
                    source.AutoAdvance ? source.SlideTimeMilliseconds : 0);
            }

            internal bool IsEquivalentTo(LegacyPptWriterTransition? other) =>
                other != null
                && EffectType == other.EffectType
                && EffectDirection == other.EffectDirection
                && Speed == other.Speed
                && ManualAdvance == other.ManualAdvance
                && AutoAdvance == other.AutoAdvance
                && SlideTimeMilliseconds == other.SlideTimeMilliseconds;
        }
    }
}
