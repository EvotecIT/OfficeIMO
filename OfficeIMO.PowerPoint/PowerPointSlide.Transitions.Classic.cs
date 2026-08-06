using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private static PowerPointSlideTransition? GetClassicTransition(Transition transition) {
            FadeTransition? fade = transition.GetFirstChild<FadeTransition>();
            if (fade != null) {
                return fade.ThroughBlack?.Value == true
                    ? PowerPointSlideTransition.FadeThroughBlack
                    : PowerPointSlideTransition.Fade;
            }

            WipeTransition? wipe = transition.GetFirstChild<WipeTransition>();
            if (wipe != null) {
                TransitionSlideDirectionValues? direction = wipe.Direction?.Value;
                if (direction == TransitionSlideDirectionValues.Up) return PowerPointSlideTransition.WipeUp;
                if (direction == TransitionSlideDirectionValues.Right) return PowerPointSlideTransition.WipeRight;
                if (direction == TransitionSlideDirectionValues.Down) return PowerPointSlideTransition.WipeDown;
                return PowerPointSlideTransition.Wipe;
            }

            BlindsTransition? blinds = transition.GetFirstChild<BlindsTransition>();
            if (blinds != null) {
                return blinds.Direction?.Value == DirectionValues.Vertical
                    ? PowerPointSlideTransition.BlindsVertical
                    : PowerPointSlideTransition.BlindsHorizontal;
            }

            CheckerTransition? checker = transition.GetFirstChild<CheckerTransition>();
            if (checker != null) {
                return checker.Direction?.Value == DirectionValues.Vertical
                    ? PowerPointSlideTransition.CheckerVertical
                    : PowerPointSlideTransition.CheckerHorizontal;
            }

            CombTransition? comb = transition.GetFirstChild<CombTransition>();
            if (comb != null) {
                return comb.Direction?.Value == DirectionValues.Vertical
                    ? PowerPointSlideTransition.CombVertical
                    : PowerPointSlideTransition.CombHorizontal;
            }

            CoverTransition? cover = transition.GetFirstChild<CoverTransition>();
            if (cover != null) {
                return MapEightDirection(cover.Direction?.Value,
                    PowerPointSlideTransition.CoverLeft, PowerPointSlideTransition.CoverUp,
                    PowerPointSlideTransition.CoverRight, PowerPointSlideTransition.CoverDown,
                    PowerPointSlideTransition.CoverLeftUp, PowerPointSlideTransition.CoverRightUp,
                    PowerPointSlideTransition.CoverLeftDown, PowerPointSlideTransition.CoverRightDown);
            }

            PullTransition? pull = transition.GetFirstChild<PullTransition>();
            if (pull != null) {
                return MapEightDirection(pull.Direction?.Value,
                    PowerPointSlideTransition.UncoverLeft, PowerPointSlideTransition.UncoverUp,
                    PowerPointSlideTransition.UncoverRight, PowerPointSlideTransition.UncoverDown,
                    PowerPointSlideTransition.UncoverLeftUp, PowerPointSlideTransition.UncoverRightUp,
                    PowerPointSlideTransition.UncoverLeftDown, PowerPointSlideTransition.UncoverRightDown);
            }

            RandomBarTransition? randomBars =
                transition.GetFirstChild<RandomBarTransition>();
            if (randomBars != null) {
                return randomBars.Direction?.Value == DirectionValues.Vertical
                    ? PowerPointSlideTransition.RandomBarsVertical
                    : PowerPointSlideTransition.RandomBarsHorizontal;
            }

            StripsTransition? strips = transition.GetFirstChild<StripsTransition>();
            if (strips != null) {
                TransitionCornerDirectionValues? direction = strips.Direction?.Value;
                if (direction == TransitionCornerDirectionValues.RightUp) return PowerPointSlideTransition.StripsRightUp;
                if (direction == TransitionCornerDirectionValues.LeftDown) return PowerPointSlideTransition.StripsLeftDown;
                if (direction == TransitionCornerDirectionValues.RightDown) return PowerPointSlideTransition.StripsRightDown;
                return PowerPointSlideTransition.StripsLeftUp;
            }

            PushTransition? push = transition.GetFirstChild<PushTransition>();
            if (push != null) {
                TransitionSlideDirectionValues? direction = push.Direction?.Value;
                if (direction == TransitionSlideDirectionValues.Up) return PowerPointSlideTransition.PushUp;
                if (direction == TransitionSlideDirectionValues.Down) return PowerPointSlideTransition.PushDown;
                if (direction == TransitionSlideDirectionValues.Right) return PowerPointSlideTransition.PushRight;
                return PowerPointSlideTransition.PushLeft;
            }

            ZoomTransition? zoom = transition.GetFirstChild<ZoomTransition>();
            if (zoom != null) {
                return zoom.Direction?.Value == TransitionInOutDirectionValues.In
                    ? PowerPointSlideTransition.BoxIn
                    : PowerPointSlideTransition.BoxOut;
            }

            SplitTransition? split = transition.GetFirstChild<SplitTransition>();
            if (split != null) {
                bool vertical = split.Orientation?.Value == DirectionValues.Vertical;
                bool inward = split.Direction?.Value == TransitionInOutDirectionValues.In;
                return vertical
                    ? inward ? PowerPointSlideTransition.SplitVerticalIn : PowerPointSlideTransition.SplitVerticalOut
                    : inward ? PowerPointSlideTransition.SplitHorizontalIn : PowerPointSlideTransition.SplitHorizontalOut;
            }

            WheelTransition? wheel = transition.GetFirstChild<WheelTransition>();
            if (wheel != null) {
                return wheel.Spokes?.Value switch {
                    1 => PowerPointSlideTransition.WheelOneSpoke,
                    2 => PowerPointSlideTransition.WheelTwoSpokes,
                    3 => PowerPointSlideTransition.WheelThreeSpokes,
                    4 => PowerPointSlideTransition.WheelFourSpokes,
                    8 => PowerPointSlideTransition.WheelEightSpokes,
                    _ => null
                };
            }

            if (transition.GetFirstChild<RandomTransition>() != null) {
                return PowerPointSlideTransition.Random;
            }
            if (transition.GetFirstChild<DissolveTransition>() != null) {
                return PowerPointSlideTransition.Dissolve;
            }
            if (transition.GetFirstChild<DiamondTransition>() != null) {
                return PowerPointSlideTransition.Diamond;
            }
            if (transition.GetFirstChild<PlusTransition>() != null) {
                return PowerPointSlideTransition.Plus;
            }
            if (transition.GetFirstChild<WedgeTransition>() != null) {
                return PowerPointSlideTransition.Wedge;
            }
            if (transition.GetFirstChild<NewsflashTransition>() != null) {
                return PowerPointSlideTransition.Newsflash;
            }
            if (transition.GetFirstChild<CircleTransition>() != null) {
                return PowerPointSlideTransition.Circle;
            }

            CutTransition? cut = transition.GetFirstChild<CutTransition>();
            if (cut != null) {
                return cut.ThroughBlack?.Value == true
                    ? PowerPointSlideTransition.CutThroughBlack
                    : PowerPointSlideTransition.Cut;
            }

            return null;
        }

        private static PowerPointSlideTransition MapEightDirection(string? direction,
            PowerPointSlideTransition left, PowerPointSlideTransition up, PowerPointSlideTransition right,
            PowerPointSlideTransition down, PowerPointSlideTransition leftUp, PowerPointSlideTransition rightUp,
            PowerPointSlideTransition leftDown, PowerPointSlideTransition rightDown) => direction switch {
                "u" => up,
                "r" => right,
                "d" => down,
                "lu" => leftUp,
                "ru" => rightUp,
                "ld" => leftDown,
                "rd" => rightDown,
                _ => left
            };

        private static OpenXmlElement? CreateClassicTransition(
            PowerPointSlideTransition transition) => transition switch {
                PowerPointSlideTransition.Fade => new FadeTransition { ThroughBlack = false },
                PowerPointSlideTransition.FadeThroughBlack => new FadeTransition { ThroughBlack = true },
                PowerPointSlideTransition.Wipe => CreateWipe(TransitionSlideDirectionValues.Left),
                PowerPointSlideTransition.WipeUp => CreateWipe(TransitionSlideDirectionValues.Up),
                PowerPointSlideTransition.WipeRight => CreateWipe(TransitionSlideDirectionValues.Right),
                PowerPointSlideTransition.WipeDown => CreateWipe(TransitionSlideDirectionValues.Down),
                PowerPointSlideTransition.BlindsVertical => new BlindsTransition { Direction = DirectionValues.Vertical },
                PowerPointSlideTransition.BlindsHorizontal => new BlindsTransition { Direction = DirectionValues.Horizontal },
                PowerPointSlideTransition.CheckerHorizontal => new CheckerTransition { Direction = DirectionValues.Horizontal },
                PowerPointSlideTransition.CheckerVertical => new CheckerTransition { Direction = DirectionValues.Vertical },
                PowerPointSlideTransition.CombHorizontal => new CombTransition { Direction = DirectionValues.Horizontal },
                PowerPointSlideTransition.CombVertical => new CombTransition { Direction = DirectionValues.Vertical },
                PowerPointSlideTransition.CoverLeft => CreateCover("l"),
                PowerPointSlideTransition.CoverUp => CreateCover("u"),
                PowerPointSlideTransition.CoverRight => CreateCover("r"),
                PowerPointSlideTransition.CoverDown => CreateCover("d"),
                PowerPointSlideTransition.CoverLeftUp => CreateCover("lu"),
                PowerPointSlideTransition.CoverRightUp => CreateCover("ru"),
                PowerPointSlideTransition.CoverLeftDown => CreateCover("ld"),
                PowerPointSlideTransition.CoverRightDown => CreateCover("rd"),
                PowerPointSlideTransition.UncoverLeft => CreatePull("l"),
                PowerPointSlideTransition.UncoverUp => CreatePull("u"),
                PowerPointSlideTransition.UncoverRight => CreatePull("r"),
                PowerPointSlideTransition.UncoverDown => CreatePull("d"),
                PowerPointSlideTransition.UncoverLeftUp => CreatePull("lu"),
                PowerPointSlideTransition.UncoverRightUp => CreatePull("ru"),
                PowerPointSlideTransition.UncoverLeftDown => CreatePull("ld"),
                PowerPointSlideTransition.UncoverRightDown => CreatePull("rd"),
                PowerPointSlideTransition.RandomBarsHorizontal => new RandomBarTransition { Direction = DirectionValues.Horizontal },
                PowerPointSlideTransition.RandomBarsVertical => new RandomBarTransition { Direction = DirectionValues.Vertical },
                PowerPointSlideTransition.StripsLeftUp => CreateStrips(TransitionCornerDirectionValues.LeftUp),
                PowerPointSlideTransition.StripsRightUp => CreateStrips(TransitionCornerDirectionValues.RightUp),
                PowerPointSlideTransition.StripsLeftDown => CreateStrips(TransitionCornerDirectionValues.LeftDown),
                PowerPointSlideTransition.StripsRightDown => CreateStrips(TransitionCornerDirectionValues.RightDown),
                PowerPointSlideTransition.PushLeft => CreatePush(TransitionSlideDirectionValues.Left),
                PowerPointSlideTransition.PushUp => CreatePush(TransitionSlideDirectionValues.Up),
                PowerPointSlideTransition.PushRight => CreatePush(TransitionSlideDirectionValues.Right),
                PowerPointSlideTransition.PushDown => CreatePush(TransitionSlideDirectionValues.Down),
                PowerPointSlideTransition.BoxOut => new ZoomTransition { Direction = TransitionInOutDirectionValues.Out },
                PowerPointSlideTransition.BoxIn => new ZoomTransition { Direction = TransitionInOutDirectionValues.In },
                PowerPointSlideTransition.SplitHorizontalOut => CreateSplit(DirectionValues.Horizontal, TransitionInOutDirectionValues.Out),
                PowerPointSlideTransition.SplitHorizontalIn => CreateSplit(DirectionValues.Horizontal, TransitionInOutDirectionValues.In),
                PowerPointSlideTransition.SplitVerticalOut => CreateSplit(DirectionValues.Vertical, TransitionInOutDirectionValues.Out),
                PowerPointSlideTransition.SplitVerticalIn => CreateSplit(DirectionValues.Vertical, TransitionInOutDirectionValues.In),
                PowerPointSlideTransition.Random => new RandomTransition(),
                PowerPointSlideTransition.Dissolve => new DissolveTransition(),
                PowerPointSlideTransition.Diamond => new DiamondTransition(),
                PowerPointSlideTransition.Plus => new PlusTransition(),
                PowerPointSlideTransition.Wedge => new WedgeTransition(),
                PowerPointSlideTransition.Newsflash => new NewsflashTransition(),
                PowerPointSlideTransition.WheelOneSpoke => new WheelTransition { Spokes = 1U },
                PowerPointSlideTransition.WheelTwoSpokes => new WheelTransition { Spokes = 2U },
                PowerPointSlideTransition.WheelThreeSpokes => new WheelTransition { Spokes = 3U },
                PowerPointSlideTransition.WheelFourSpokes => new WheelTransition { Spokes = 4U },
                PowerPointSlideTransition.WheelEightSpokes => new WheelTransition { Spokes = 8U },
                PowerPointSlideTransition.Circle => new CircleTransition(),
                PowerPointSlideTransition.Cut => new CutTransition { ThroughBlack = false },
                PowerPointSlideTransition.CutThroughBlack => new CutTransition { ThroughBlack = true },
                _ => null
            };

        private static WipeTransition CreateWipe(
            TransitionSlideDirectionValues direction) =>
            new() { Direction = direction };

        private static PushTransition CreatePush(
            TransitionSlideDirectionValues direction) =>
            new() { Direction = direction };

        private static CoverTransition CreateCover(string direction) =>
            new() { Direction = direction };

        private static PullTransition CreatePull(string direction) =>
            new() { Direction = direction };

        private static StripsTransition CreateStrips(
            TransitionCornerDirectionValues direction) =>
            new() { Direction = direction };

        private static SplitTransition CreateSplit(DirectionValues orientation,
            TransitionInOutDirectionValues direction) =>
            new() { Orientation = orientation, Direction = direction };
    }
}
