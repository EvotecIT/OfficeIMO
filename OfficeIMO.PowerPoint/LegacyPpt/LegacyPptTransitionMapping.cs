using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    /// <summary>
    /// Maps the shared editable transition surface to exact PowerPoint 97-2003
    /// SlideShowSlideInfoAtom effect and direction values.
    /// </summary>
    internal static class LegacyPptTransitionMapping {
        internal static bool TryGetBinary(SlideTransition transition,
            out byte effectType, out byte effectDirection) {
            (byte EffectType, byte EffectDirection) mapping = transition switch {
                SlideTransition.Cut => (0, 0),
                SlideTransition.CutThroughBlack => (0, 1),
                SlideTransition.Random => (1, 0),
                SlideTransition.BlindsVertical => (2, 0),
                SlideTransition.BlindsHorizontal => (2, 1),
                SlideTransition.CheckerHorizontal => (3, 0),
                SlideTransition.CheckerVertical => (3, 1),
                SlideTransition.CoverLeft => (4, 0),
                SlideTransition.CoverUp => (4, 1),
                SlideTransition.CoverRight => (4, 2),
                SlideTransition.CoverDown => (4, 3),
                SlideTransition.CoverLeftUp => (4, 4),
                SlideTransition.CoverRightUp => (4, 5),
                SlideTransition.CoverLeftDown => (4, 6),
                SlideTransition.CoverRightDown => (4, 7),
                SlideTransition.Dissolve => (5, 0),
                SlideTransition.FadeThroughBlack => (6, 0),
                SlideTransition.UncoverLeft => (7, 0),
                SlideTransition.UncoverUp => (7, 1),
                SlideTransition.UncoverRight => (7, 2),
                SlideTransition.UncoverDown => (7, 3),
                SlideTransition.UncoverLeftUp => (7, 4),
                SlideTransition.UncoverRightUp => (7, 5),
                SlideTransition.UncoverLeftDown => (7, 6),
                SlideTransition.UncoverRightDown => (7, 7),
                SlideTransition.RandomBarsHorizontal => (8, 0),
                SlideTransition.RandomBarsVertical => (8, 1),
                SlideTransition.StripsLeftUp => (9, 4),
                SlideTransition.StripsRightUp => (9, 5),
                SlideTransition.StripsLeftDown => (9, 6),
                SlideTransition.StripsRightDown => (9, 7),
                SlideTransition.Wipe => (10, 0),
                SlideTransition.WipeUp => (10, 1),
                SlideTransition.WipeRight => (10, 2),
                SlideTransition.WipeDown => (10, 3),
                SlideTransition.BoxOut => (11, 0),
                SlideTransition.BoxIn => (11, 1),
                SlideTransition.SplitHorizontalOut => (13, 0),
                SlideTransition.SplitHorizontalIn => (13, 1),
                SlideTransition.SplitVerticalOut => (13, 2),
                SlideTransition.SplitVerticalIn => (13, 3),
                SlideTransition.Diamond => (17, 0),
                SlideTransition.Plus => (18, 0),
                SlideTransition.Wedge => (19, 0),
                SlideTransition.PushLeft => (20, 0),
                SlideTransition.PushUp => (20, 1),
                SlideTransition.PushRight => (20, 2),
                SlideTransition.PushDown => (20, 3),
                SlideTransition.CombHorizontal => (21, 0),
                SlideTransition.CombVertical => (21, 1),
                SlideTransition.Newsflash => (22, 0),
                SlideTransition.Fade => (23, 0),
                SlideTransition.WheelOneSpoke => (26, 1),
                SlideTransition.WheelTwoSpokes => (26, 2),
                SlideTransition.WheelThreeSpokes => (26, 3),
                SlideTransition.WheelFourSpokes => (26, 4),
                SlideTransition.WheelEightSpokes => (26, 8),
                SlideTransition.Circle => (27, 0),
                _ => (byte.MaxValue, byte.MaxValue)
            };
            effectType = mapping.EffectType;
            effectDirection = mapping.EffectDirection;
            return effectType != byte.MaxValue;
        }

        internal static SlideTransition? ToSlideTransition(
            LegacyPptTransition source) => source.Effect switch {
                LegacyPptTransitionEffect.Cut => source.EffectDirection switch {
                    0 => SlideTransition.Cut,
                    1 => SlideTransition.CutThroughBlack,
                    _ => null
                },
                LegacyPptTransitionEffect.Random => SlideTransition.Random,
                LegacyPptTransitionEffect.Blinds => source.EffectDirection switch {
                    0 => SlideTransition.BlindsVertical,
                    1 => SlideTransition.BlindsHorizontal,
                    _ => null
                },
                LegacyPptTransitionEffect.Checker => source.EffectDirection switch {
                    0 => SlideTransition.CheckerHorizontal,
                    1 => SlideTransition.CheckerVertical,
                    _ => null
                },
                LegacyPptTransitionEffect.Cover => MapEightDirection(source.EffectDirection,
                    SlideTransition.CoverLeft, SlideTransition.CoverUp,
                    SlideTransition.CoverRight, SlideTransition.CoverDown,
                    SlideTransition.CoverLeftUp, SlideTransition.CoverRightUp,
                    SlideTransition.CoverLeftDown, SlideTransition.CoverRightDown),
                LegacyPptTransitionEffect.Dissolve => source.EffectDirection == 0
                    ? SlideTransition.Dissolve : null,
                LegacyPptTransitionEffect.Fade => source.EffectDirection == 0
                    ? SlideTransition.FadeThroughBlack : null,
                LegacyPptTransitionEffect.Uncover => MapEightDirection(source.EffectDirection,
                    SlideTransition.UncoverLeft, SlideTransition.UncoverUp,
                    SlideTransition.UncoverRight, SlideTransition.UncoverDown,
                    SlideTransition.UncoverLeftUp, SlideTransition.UncoverRightUp,
                    SlideTransition.UncoverLeftDown, SlideTransition.UncoverRightDown),
                LegacyPptTransitionEffect.RandomBars => source.EffectDirection switch {
                    0 => SlideTransition.RandomBarsHorizontal,
                    1 => SlideTransition.RandomBarsVertical,
                    _ => null
                },
                LegacyPptTransitionEffect.Strips => source.EffectDirection switch {
                    4 => SlideTransition.StripsLeftUp,
                    5 => SlideTransition.StripsRightUp,
                    6 => SlideTransition.StripsLeftDown,
                    7 => SlideTransition.StripsRightDown,
                    _ => null
                },
                LegacyPptTransitionEffect.Wipe => source.EffectDirection switch {
                    0 => SlideTransition.Wipe,
                    1 => SlideTransition.WipeUp,
                    2 => SlideTransition.WipeRight,
                    3 => SlideTransition.WipeDown,
                    _ => null
                },
                LegacyPptTransitionEffect.Box => source.EffectDirection switch {
                    0 => SlideTransition.BoxOut,
                    1 => SlideTransition.BoxIn,
                    _ => null
                },
                LegacyPptTransitionEffect.Split => source.EffectDirection switch {
                    0 => SlideTransition.SplitHorizontalOut,
                    1 => SlideTransition.SplitHorizontalIn,
                    2 => SlideTransition.SplitVerticalOut,
                    3 => SlideTransition.SplitVerticalIn,
                    _ => null
                },
                LegacyPptTransitionEffect.Diamond => source.EffectDirection == 0
                    ? SlideTransition.Diamond : null,
                LegacyPptTransitionEffect.Plus => source.EffectDirection == 0
                    ? SlideTransition.Plus : null,
                LegacyPptTransitionEffect.Wedge => source.EffectDirection == 0
                    ? SlideTransition.Wedge : null,
                LegacyPptTransitionEffect.Push => source.EffectDirection switch {
                    0 => SlideTransition.PushLeft,
                    1 => SlideTransition.PushUp,
                    2 => SlideTransition.PushRight,
                    3 => SlideTransition.PushDown,
                    _ => null
                },
                LegacyPptTransitionEffect.Comb => source.EffectDirection switch {
                    0 => SlideTransition.CombHorizontal,
                    1 => SlideTransition.CombVertical,
                    _ => null
                },
                LegacyPptTransitionEffect.Newsflash => source.EffectDirection == 0
                    ? SlideTransition.Newsflash : null,
                LegacyPptTransitionEffect.AlphaFade => source.EffectDirection == 0
                    ? SlideTransition.Fade : null,
                LegacyPptTransitionEffect.Wheel => source.EffectDirection switch {
                    1 => SlideTransition.WheelOneSpoke,
                    2 => SlideTransition.WheelTwoSpokes,
                    3 => SlideTransition.WheelThreeSpokes,
                    4 => SlideTransition.WheelFourSpokes,
                    8 => SlideTransition.WheelEightSpokes,
                    _ => null
                },
                LegacyPptTransitionEffect.Circle => source.EffectDirection == 0
                    ? SlideTransition.Circle : null,
                _ => null
            };

        private static SlideTransition? MapEightDirection(byte direction,
            SlideTransition left, SlideTransition up, SlideTransition right,
            SlideTransition down, SlideTransition leftUp, SlideTransition rightUp,
            SlideTransition leftDown, SlideTransition rightDown) => direction switch {
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
