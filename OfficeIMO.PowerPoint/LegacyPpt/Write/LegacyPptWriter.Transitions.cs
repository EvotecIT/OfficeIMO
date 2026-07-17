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
            IReadOnlyList<DocumentFormat.OpenXml.Presentation.Transition?>
                transitionBranches = slide.GetTransitionBranches();
            LegacyPptWriterTransitionBranch? firstBranch = null;
            foreach (DocumentFormat.OpenXml.Presentation.Transition?
                     transitionElement in transitionBranches) {
                if (!TryReadTransitionBranch(transitionElement,
                        slide.SlidePart, soundCatalog,
                        out LegacyPptWriterTransitionBranch? branch,
                        out reason)) {
                    return false;
                }
                if (firstBranch != null
                    && !firstBranch.IsEquivalentTo(branch!)) {
                    reason = "AlternateContent transition branches contain inconsistent effects, timing, advance settings, or sound actions.";
                    return false;
                }
                firstBranch ??= branch;
            }
            if (firstBranch?.HasData != true) return true;
            transition = firstBranch.Transition;
            return true;
        }

        private static bool TryReadTransitionBranch(
            DocumentFormat.OpenXml.Presentation.Transition? transition,
            DocumentFormat.OpenXml.Packaging.SlidePart ownerPart,
            LegacyPptWriterSoundCatalog soundCatalog,
            out LegacyPptWriterTransitionBranch? branch,
            out string? reason) {
            branch = null;
            reason = null;
            if (transition == null) {
                branch = new LegacyPptWriterTransitionBranch(
                    hasData: false, hasEffect: false, soundKind: "none",
                    soundBuiltIn: false,
                    new LegacyPptWriterTransition(effectType: 0,
                        effectDirection: 0, speed: 1,
                        manualAdvance: true, autoAdvance: false,
                        slideTimeMilliseconds: 0));
                return true;
            }
            SlideTransition effect = PowerPointSlide.GetTransitionValue(
                transition);
            bool hasEffect = effect != SlideTransition.None;
            bool hasEffectMarkup = transition.ChildElements.Any(child =>
                child is not DocumentFormat.OpenXml.Presentation.SoundAction
                && !string.Equals(child.LocalName, "extLst",
                    StringComparison.Ordinal));
            byte effectType = 0;
            byte effectDirection = 0;
            if (hasEffect && !LegacyPptTransitionMapping.TryGetBinary(effect,
                    out effectType, out effectDirection)) {
                reason = $"The {effect} transition has no PowerPoint 97-2003 representation.";
                return false;
            }
            if (hasEffectMarkup && !hasEffect) {
                reason = "A transition branch contains an effect with no PowerPoint 97-2003 representation.";
                return false;
            }

            if (!TryReadTransitionSoundAction(transition, ownerPart,
                    soundCatalog,
                    out LegacyPptWriterSound? resolvedSound,
                    out (string Kind, uint? SoundId, bool BuiltIn,
                        bool Loop) soundSignature,
                    out reason)) {
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
                if (resolvedSound == null) {
                    reason = "A transition start-sound action must contain exactly one representable embedded sound.";
                    return false;
                }
                soundId = resolvedSound.Id;
                playSound = true;
                loopSound = soundSignature.Loop;
            }

            DocumentFormat.OpenXml.Presentation.TransitionSpeedValues?
                speedValue = transition.Speed?.Value;
            byte speed = speedValue
                         == DocumentFormat.OpenXml.Presentation
                             .TransitionSpeedValues.Slow
                ? (byte)0
                : speedValue == DocumentFormat.OpenXml.Presentation
                    .TransitionSpeedValues.Fast
                    ? (byte)2
                    : (byte)1;
            string? durationText = transition.Duration?.Value;
            if (!string.IsNullOrWhiteSpace(durationText)) {
                if (!uint.TryParse(durationText,
                        System.Globalization.NumberStyles.None,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out uint durationMilliseconds)) {
                    reason = "A transition branch contains an invalid duration.";
                    return false;
                }
                byte? durationSpeed = durationMilliseconds switch {
                    750 => 0,
                    500 => 1,
                    250 => 2,
                    _ => null
                };
                if (!durationSpeed.HasValue) {
                    reason = "Binary PowerPoint transition duration must be exactly 0.75, 0.5, or 0.25 seconds.";
                    return false;
                }
                if (transition.Speed?.Value != null
                    && speed != durationSpeed.Value) {
                    reason = "The requested transition speed and duration describe different binary speed values.";
                    return false;
                }
                speed = durationSpeed.Value;
            }

            int slideTime = 0;
            string? advanceAfterText = transition.AdvanceAfterTime?.Value;
            bool autoAdvance = !string.IsNullOrWhiteSpace(
                advanceAfterText);
            if (autoAdvance) {
                if (!uint.TryParse(advanceAfterText,
                        System.Globalization.NumberStyles.None,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out uint advanceMilliseconds)) {
                    reason = "A transition branch contains an invalid automatic-advance time.";
                    return false;
                }
                if (advanceMilliseconds > 86399000U) {
                    reason = "Binary PowerPoint automatic advance must be between 0 and 86,399 seconds.";
                    return false;
                }
                slideTime = checked((int)advanceMilliseconds);
            }
            bool hasSettings = transition.Speed?.Value != null
                || !string.IsNullOrWhiteSpace(durationText)
                || transition.AdvanceOnClick?.Value != null
                || autoAdvance;
            var projected = new LegacyPptWriterTransition(effectType,
                effectDirection, speed,
                transition.AdvanceOnClick?.Value != false, autoAdvance,
                slideTime, soundId, playSound, loopSound, stopSound);
            branch = new LegacyPptWriterTransitionBranch(
                hasEffect || hasSettings
                    || !string.Equals(soundSignature.Kind, "none",
                        StringComparison.Ordinal),
                hasEffect, soundSignature.Kind, soundSignature.BuiltIn,
                projected);
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

        private sealed class LegacyPptWriterTransitionBranch {
            internal LegacyPptWriterTransitionBranch(bool hasData,
                bool hasEffect, string soundKind, bool soundBuiltIn,
                LegacyPptWriterTransition transition) {
                HasData = hasData;
                HasEffect = hasEffect;
                SoundKind = soundKind;
                SoundBuiltIn = soundBuiltIn;
                Transition = transition;
            }

            internal bool HasData { get; }
            internal bool HasEffect { get; }
            internal string SoundKind { get; }
            internal bool SoundBuiltIn { get; }
            internal LegacyPptWriterTransition Transition { get; }

            internal bool IsEquivalentTo(
                LegacyPptWriterTransitionBranch other) =>
                other != null
                && HasEffect == other.HasEffect
                && string.Equals(SoundKind, other.SoundKind,
                    StringComparison.Ordinal)
                && SoundBuiltIn == other.SoundBuiltIn
                && Transition.EffectType == other.Transition.EffectType
                && Transition.EffectDirection
                    == other.Transition.EffectDirection
                && Transition.Speed == other.Transition.Speed
                && Transition.ManualAdvance
                    == other.Transition.ManualAdvance
                && Transition.AutoAdvance == other.Transition.AutoAdvance
                && Transition.SlideTimeMilliseconds
                    == other.Transition.SlideTimeMilliseconds
                && Transition.SoundId == other.Transition.SoundId
                && Transition.PlaySound == other.Transition.PlaySound
                && Transition.LoopSound == other.Transition.LoopSound
                && Transition.StopSound == other.Transition.StopSound;
        }

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
