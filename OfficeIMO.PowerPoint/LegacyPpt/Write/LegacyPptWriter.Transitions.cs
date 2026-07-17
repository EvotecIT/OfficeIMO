namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal static bool TryReadTransition(PowerPointSlide slide,
            out LegacyPptWriterTransition? transition, out string? reason) {
            return TryReadTransition(slide, new LegacyPptWriterSoundCatalog(),
                out transition, out reason);
        }

        internal static bool TryReadTransition(PowerPointSlide slide,
            LegacyPptWriterSoundCatalog soundCatalog,
            out LegacyPptWriterTransition? transition, out string? reason) {
            transition = null;
            reason = null;
            IReadOnlyList<DocumentFormat.OpenXml.Presentation.Transition>
                transitionElements = slide.GetTransitionElements();
            LegacyPptWriterSound? transitionSound = null;
            (string Kind, uint? SoundId, bool BuiltIn, bool Loop)
                soundSignature = default;
            bool hasSoundSignature = false;
            foreach (DocumentFormat.OpenXml.Presentation.Transition
                     transitionElement in transitionElements) {
                if (!TryReadTransitionSoundAction(transitionElement,
                        slide.SlidePart, soundCatalog,
                        out LegacyPptWriterSound? branchSound,
                        out (string Kind, uint? SoundId, bool BuiltIn,
                            bool Loop) branchSignature,
                        out reason)) {
                    return false;
                }
                if (hasSoundSignature
                    && !soundSignature.Equals(branchSignature)) {
                    reason = "AlternateContent transition branches contain inconsistent sound actions.";
                    return false;
                }
                if (!hasSoundSignature) {
                    transitionSound = branchSound;
                    soundSignature = branchSignature;
                    hasSoundSignature = true;
                }
            }
            bool hasTransitionEffect = slide.Transition != SlideTransition.None;
            bool hasSoundAction = hasSoundSignature
                && !string.Equals(soundSignature.Kind, "none",
                    StringComparison.Ordinal);
            if (!hasTransitionEffect && !hasSoundAction) return true;

            byte effectType = 0;
            byte effectDirection = 0;
            if (hasTransitionEffect && !LegacyPptTransitionMapping.TryGetBinary(slide.Transition,
                    out effectType, out effectDirection)) {
                reason = $"The {slide.Transition} transition has no PowerPoint 97-2003 representation.";
                return false;
            }
            uint soundId = 0;
            bool playSound = false;
            bool loopSound = false;
            bool stopSound = false;
            if (string.Equals(soundSignature.Kind, "end",
                    StringComparison.Ordinal)) {
                stopSound = true;
            } else if (string.Equals(soundSignature.Kind, "start",
                           StringComparison.Ordinal)) {
                if (transitionSound == null) {
                    reason = "A transition start-sound action must contain exactly one representable embedded sound.";
                    return false;
                }
                soundId = transitionSound.Id;
                playSound = true;
                loopSound = soundSignature.Loop;
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
                slideTime, soundId, playSound, loopSound, stopSound);
            return true;
        }

        private static bool TryReadTransitionSoundAction(
            DocumentFormat.OpenXml.Presentation.Transition transition,
            DocumentFormat.OpenXml.Packaging.SlidePart ownerPart,
            LegacyPptWriterSoundCatalog soundCatalog,
            out LegacyPptWriterSound? resolvedSound,
            out (string Kind, uint? SoundId, bool BuiltIn, bool Loop)
                signature,
            out string? reason) {
            DocumentFormat.OpenXml.Presentation.SoundAction[] actions =
                transition.Elements<
                    DocumentFormat.OpenXml.Presentation.SoundAction>()
                    .ToArray();
            resolvedSound = null;
            signature = ("none", null, false, false);
            reason = null;
            if (actions.Length == 0) return true;
            if (actions.Length != 1) {
                reason = "A transition branch must contain at most one sound action.";
                return false;
            }
            DocumentFormat.OpenXml.Presentation.SoundAction soundAction =
                actions[0];
            DocumentFormat.OpenXml.Presentation.StartSoundAction[] starts =
                soundAction.Elements<
                    DocumentFormat.OpenXml.Presentation.StartSoundAction>()
                    .ToArray();
            DocumentFormat.OpenXml.Presentation.EndSoundAction[] ends =
                soundAction.Elements<
                    DocumentFormat.OpenXml.Presentation.EndSoundAction>()
                    .ToArray();
            if (soundAction.ChildElements.Count != 1
                || starts.Length + ends.Length != 1) {
                reason = "A transition sound action must contain exactly one start-sound or end-sound action.";
                return false;
            }
            if (ends.Length == 1) {
                signature = ("end", null, false, false);
                return true;
            }
            DocumentFormat.OpenXml.Presentation.Sound[] sounds = starts[0]
                .Elements<DocumentFormat.OpenXml.Presentation.Sound>()
                .ToArray();
            if (starts[0].ChildElements.Count != 1
                || sounds.Length != 1) {
                reason = "A transition start-sound action must contain exactly one embedded sound.";
                return false;
            }
            if (!soundCatalog.TryGetOrAdd(ownerPart, sounds[0],
                    out resolvedSound, out reason)) {
                reason ??= "A transition start-sound action must contain exactly one representable embedded sound.";
                return false;
            }
            signature = ("start", resolvedSound!.Id,
                sounds[0].BuiltIn?.Value == true,
                starts[0].Loop?.Value == true);
            return true;
        }

        private static bool NearlyEqual(double left, double right) =>
            Math.Abs(left - right) < 0.0005;

        internal static byte[] PatchSlideShowInfo(byte[] record,
            PowerPointSlide slide,
            LegacyPptWriterSoundCatalog? soundCatalog = null) {
            if (record.Length < 24) {
                throw new InvalidDataException(
                    "The slide-show information atom is too short.");
            }
            if (!TryReadTransition(slide,
                    soundCatalog ?? new LegacyPptWriterSoundCatalog(),
                    out LegacyPptWriterTransition? transition,
                    out string? reason)) {
                throw new NotSupportedException(reason);
            }
            PatchSlideShowInfoPayload(record, slide.Hidden, transition);
            return record;
        }

        internal static byte[] BuildSlideShowInfoRecord(PowerPointSlide slide,
            LegacyPptWriterSoundCatalog? soundCatalog = null) {
            if (!TryReadTransition(slide,
                    soundCatalog ?? new LegacyPptWriterSoundCatalog(),
                    out LegacyPptWriterTransition? transition,
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
            WriteUInt32(bytes, payloadOffset + 4, transition?.SoundId ?? 0);
            bytes[payloadOffset + 8] = transition?.EffectDirection ?? 0;
            bytes[payloadOffset + 9] = transition?.EffectType ?? 0;
            ushort flags = ReadUInt16(bytes, payloadOffset + 10);
            flags = unchecked((ushort)(flags & ~(0x0001 | 0x0004 | 0x0010
                | 0x0040 | 0x0100 | 0x0400)));
            if (hidden) flags |= 0x0004;
            if (transition?.ManualAdvance != false) flags |= 0x0001;
            if (transition?.AutoAdvance == true) flags |= 0x0400;
            if (transition?.PlaySound == true) flags |= 0x0010;
            if (transition?.LoopSound == true) flags |= 0x0040;
            if (transition?.StopSound == true) flags |= 0x0100;
            WriteUInt16(bytes, payloadOffset + 10, flags);
            bytes[payloadOffset + 12] = transition?.Speed ?? 1;
        }

        internal sealed class LegacyPptWriterTransition {
            internal LegacyPptWriterTransition(byte effectType, byte effectDirection,
                byte speed, bool manualAdvance, bool autoAdvance,
                int slideTimeMilliseconds, uint soundId = 0,
                bool playSound = false, bool loopSound = false,
                bool stopSound = false) {
                EffectType = effectType;
                EffectDirection = effectDirection;
                Speed = speed;
                ManualAdvance = manualAdvance;
                AutoAdvance = autoAdvance;
                SlideTimeMilliseconds = slideTimeMilliseconds;
                SoundId = soundId;
                PlaySound = playSound;
                LoopSound = loopSound;
                StopSound = stopSound;
            }

            internal byte EffectType { get; }
            internal byte EffectDirection { get; }
            internal byte Speed { get; }
            internal bool ManualAdvance { get; }
            internal bool AutoAdvance { get; }
            internal int SlideTimeMilliseconds { get; }
            internal uint SoundId { get; }
            internal bool PlaySound { get; }
            internal bool LoopSound { get; }
            internal bool StopSound { get; }

            internal static LegacyPptWriterTransition? FromLegacyProjection(
                OfficeIMO.PowerPoint.LegacyPpt.Model.LegacyPptTransition? source) {
                if (source == null) return null;
                SlideTransition? projected =
                    LegacyPptTransitionMapping.ToSlideTransition(source);
                if (!projected.HasValue
                    || !LegacyPptTransitionMapping.TryGetBinary(projected.Value,
                        out byte effectType, out byte effectDirection)) return null;
                byte speed = source.Speed <= 2 ? source.Speed : (byte)1;
                return new LegacyPptWriterTransition(effectType, effectDirection,
                    speed, source.ManualAdvance, source.AutoAdvance,
                    source.AutoAdvance ? source.SlideTimeMilliseconds : 0,
                    source.SoundId, source.PlaySound, source.LoopSound,
                    source.StopSound && !source.PlaySound);
            }

            internal bool IsEquivalentTo(LegacyPptWriterTransition? other) =>
                other != null
                && EffectType == other.EffectType
                && EffectDirection == other.EffectDirection
                && Speed == other.Speed
                && ManualAdvance == other.ManualAdvance
                && AutoAdvance == other.AutoAdvance
                && SlideTimeMilliseconds == other.SlideTimeMilliseconds
                && SoundId == other.SoundId
                && PlaySound == other.PlaySound
                && LoopSound == other.LoopSound
                && StopSound == other.StopSound;
        }
    }
}
