using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    /// <summary>
    /// Maps the shared editable transition surface to exact PowerPoint 97-2003
    /// SlideShowSlideInfoAtom effect and direction values.
    /// </summary>
    internal static class LegacyPptTransitionMapping {
        internal static bool TryGetBinary(PowerPointSlideTransition transition,
            out byte effectType, out byte effectDirection) {
            (byte EffectType, byte EffectDirection) mapping = transition switch {
                PowerPointSlideTransition.Cut => (0, 0),
                PowerPointSlideTransition.CutThroughBlack => (0, 1),
                PowerPointSlideTransition.Random => (1, 0),
                PowerPointSlideTransition.BlindsVertical => (2, 0),
                PowerPointSlideTransition.BlindsHorizontal => (2, 1),
                PowerPointSlideTransition.CheckerHorizontal => (3, 0),
                PowerPointSlideTransition.CheckerVertical => (3, 1),
                PowerPointSlideTransition.CoverLeft => (4, 0),
                PowerPointSlideTransition.CoverUp => (4, 1),
                PowerPointSlideTransition.CoverRight => (4, 2),
                PowerPointSlideTransition.CoverDown => (4, 3),
                PowerPointSlideTransition.CoverLeftUp => (4, 4),
                PowerPointSlideTransition.CoverRightUp => (4, 5),
                PowerPointSlideTransition.CoverLeftDown => (4, 6),
                PowerPointSlideTransition.CoverRightDown => (4, 7),
                PowerPointSlideTransition.Dissolve => (5, 0),
                PowerPointSlideTransition.FadeThroughBlack => (6, 0),
                PowerPointSlideTransition.UncoverLeft => (7, 0),
                PowerPointSlideTransition.UncoverUp => (7, 1),
                PowerPointSlideTransition.UncoverRight => (7, 2),
                PowerPointSlideTransition.UncoverDown => (7, 3),
                PowerPointSlideTransition.UncoverLeftUp => (7, 4),
                PowerPointSlideTransition.UncoverRightUp => (7, 5),
                PowerPointSlideTransition.UncoverLeftDown => (7, 6),
                PowerPointSlideTransition.UncoverRightDown => (7, 7),
                PowerPointSlideTransition.RandomBarsHorizontal => (8, 0),
                PowerPointSlideTransition.RandomBarsVertical => (8, 1),
                PowerPointSlideTransition.StripsLeftUp => (9, 4),
                PowerPointSlideTransition.StripsRightUp => (9, 5),
                PowerPointSlideTransition.StripsLeftDown => (9, 6),
                PowerPointSlideTransition.StripsRightDown => (9, 7),
                PowerPointSlideTransition.Wipe => (10, 0),
                PowerPointSlideTransition.WipeUp => (10, 1),
                PowerPointSlideTransition.WipeRight => (10, 2),
                PowerPointSlideTransition.WipeDown => (10, 3),
                PowerPointSlideTransition.BoxOut => (11, 0),
                PowerPointSlideTransition.BoxIn => (11, 1),
                PowerPointSlideTransition.SplitHorizontalOut => (13, 0),
                PowerPointSlideTransition.SplitHorizontalIn => (13, 1),
                PowerPointSlideTransition.SplitVerticalOut => (13, 2),
                PowerPointSlideTransition.SplitVerticalIn => (13, 3),
                PowerPointSlideTransition.Diamond => (17, 0),
                PowerPointSlideTransition.Plus => (18, 0),
                PowerPointSlideTransition.Wedge => (19, 0),
                PowerPointSlideTransition.PushLeft => (20, 0),
                PowerPointSlideTransition.PushUp => (20, 1),
                PowerPointSlideTransition.PushRight => (20, 2),
                PowerPointSlideTransition.PushDown => (20, 3),
                PowerPointSlideTransition.CombHorizontal => (21, 0),
                PowerPointSlideTransition.CombVertical => (21, 1),
                PowerPointSlideTransition.Newsflash => (22, 0),
                PowerPointSlideTransition.Fade => (23, 0),
                PowerPointSlideTransition.WheelOneSpoke => (26, 1),
                PowerPointSlideTransition.WheelTwoSpokes => (26, 2),
                PowerPointSlideTransition.WheelThreeSpokes => (26, 3),
                PowerPointSlideTransition.WheelFourSpokes => (26, 4),
                PowerPointSlideTransition.WheelEightSpokes => (26, 8),
                PowerPointSlideTransition.Circle => (27, 0),
                _ => (byte.MaxValue, byte.MaxValue)
            };
            effectType = mapping.EffectType;
            effectDirection = mapping.EffectDirection;
            return effectType != byte.MaxValue;
        }

        internal static PowerPointSlideTransition? ToSlideTransition(
            LegacyPptTransition source) => source.Effect switch {
                LegacyPptTransitionEffect.Cut => source.EffectDirection switch {
                    0 => PowerPointSlideTransition.Cut,
                    1 => PowerPointSlideTransition.CutThroughBlack,
                    _ => null
                },
                LegacyPptTransitionEffect.Random => PowerPointSlideTransition.Random,
                LegacyPptTransitionEffect.Blinds => source.EffectDirection switch {
                    0 => PowerPointSlideTransition.BlindsVertical,
                    1 => PowerPointSlideTransition.BlindsHorizontal,
                    _ => null
                },
                LegacyPptTransitionEffect.Checker => source.EffectDirection switch {
                    0 => PowerPointSlideTransition.CheckerHorizontal,
                    1 => PowerPointSlideTransition.CheckerVertical,
                    _ => null
                },
                LegacyPptTransitionEffect.Cover => MapEightDirection(source.EffectDirection,
                    PowerPointSlideTransition.CoverLeft, PowerPointSlideTransition.CoverUp,
                    PowerPointSlideTransition.CoverRight, PowerPointSlideTransition.CoverDown,
                    PowerPointSlideTransition.CoverLeftUp, PowerPointSlideTransition.CoverRightUp,
                    PowerPointSlideTransition.CoverLeftDown, PowerPointSlideTransition.CoverRightDown),
                LegacyPptTransitionEffect.Dissolve => source.EffectDirection == 0
                    ? PowerPointSlideTransition.Dissolve : null,
                LegacyPptTransitionEffect.Fade => source.EffectDirection == 0
                    ? PowerPointSlideTransition.FadeThroughBlack : null,
                LegacyPptTransitionEffect.Uncover => MapEightDirection(source.EffectDirection,
                    PowerPointSlideTransition.UncoverLeft, PowerPointSlideTransition.UncoverUp,
                    PowerPointSlideTransition.UncoverRight, PowerPointSlideTransition.UncoverDown,
                    PowerPointSlideTransition.UncoverLeftUp, PowerPointSlideTransition.UncoverRightUp,
                    PowerPointSlideTransition.UncoverLeftDown, PowerPointSlideTransition.UncoverRightDown),
                LegacyPptTransitionEffect.RandomBars => source.EffectDirection switch {
                    0 => PowerPointSlideTransition.RandomBarsHorizontal,
                    1 => PowerPointSlideTransition.RandomBarsVertical,
                    _ => null
                },
                LegacyPptTransitionEffect.Strips => source.EffectDirection switch {
                    4 => PowerPointSlideTransition.StripsLeftUp,
                    5 => PowerPointSlideTransition.StripsRightUp,
                    6 => PowerPointSlideTransition.StripsLeftDown,
                    7 => PowerPointSlideTransition.StripsRightDown,
                    _ => null
                },
                LegacyPptTransitionEffect.Wipe => source.EffectDirection switch {
                    0 => PowerPointSlideTransition.Wipe,
                    1 => PowerPointSlideTransition.WipeUp,
                    2 => PowerPointSlideTransition.WipeRight,
                    3 => PowerPointSlideTransition.WipeDown,
                    _ => null
                },
                LegacyPptTransitionEffect.Box => source.EffectDirection switch {
                    0 => PowerPointSlideTransition.BoxOut,
                    1 => PowerPointSlideTransition.BoxIn,
                    _ => null
                },
                LegacyPptTransitionEffect.Split => source.EffectDirection switch {
                    0 => PowerPointSlideTransition.SplitHorizontalOut,
                    1 => PowerPointSlideTransition.SplitHorizontalIn,
                    2 => PowerPointSlideTransition.SplitVerticalOut,
                    3 => PowerPointSlideTransition.SplitVerticalIn,
                    _ => null
                },
                LegacyPptTransitionEffect.Diamond => source.EffectDirection == 0
                    ? PowerPointSlideTransition.Diamond : null,
                LegacyPptTransitionEffect.Plus => source.EffectDirection == 0
                    ? PowerPointSlideTransition.Plus : null,
                LegacyPptTransitionEffect.Wedge => source.EffectDirection == 0
                    ? PowerPointSlideTransition.Wedge : null,
                LegacyPptTransitionEffect.Push => source.EffectDirection switch {
                    0 => PowerPointSlideTransition.PushLeft,
                    1 => PowerPointSlideTransition.PushUp,
                    2 => PowerPointSlideTransition.PushRight,
                    3 => PowerPointSlideTransition.PushDown,
                    _ => null
                },
                LegacyPptTransitionEffect.Comb => source.EffectDirection switch {
                    0 => PowerPointSlideTransition.CombHorizontal,
                    1 => PowerPointSlideTransition.CombVertical,
                    _ => null
                },
                LegacyPptTransitionEffect.Newsflash => source.EffectDirection == 0
                    ? PowerPointSlideTransition.Newsflash : null,
                LegacyPptTransitionEffect.AlphaFade => source.EffectDirection == 0
                    ? PowerPointSlideTransition.Fade : null,
                LegacyPptTransitionEffect.Wheel => source.EffectDirection switch {
                    1 => PowerPointSlideTransition.WheelOneSpoke,
                    2 => PowerPointSlideTransition.WheelTwoSpokes,
                    3 => PowerPointSlideTransition.WheelThreeSpokes,
                    4 => PowerPointSlideTransition.WheelFourSpokes,
                    8 => PowerPointSlideTransition.WheelEightSpokes,
                    _ => null
                },
                LegacyPptTransitionEffect.Circle => source.EffectDirection == 0
                    ? PowerPointSlideTransition.Circle : null,
                _ => null
            };

        private static PowerPointSlideTransition? MapEightDirection(byte direction,
            PowerPointSlideTransition left, PowerPointSlideTransition up, PowerPointSlideTransition right,
            PowerPointSlideTransition down, PowerPointSlideTransition leftUp, PowerPointSlideTransition rightUp,
            PowerPointSlideTransition leftDown, PowerPointSlideTransition rightDown) => direction switch {
                0 => left,
                1 => up,
                2 => right,
                3 => down,
                4 => leftUp,
                5 => rightUp,
                6 => leftDown,
                7 => rightDown,
                _ => null
            };
    }
}
