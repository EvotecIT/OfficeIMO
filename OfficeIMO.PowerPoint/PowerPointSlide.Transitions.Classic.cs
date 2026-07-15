using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private static SlideTransition? GetClassicTransition(Transition transition) {
            FadeTransition? fade = transition.GetFirstChild<FadeTransition>();
            if (fade != null) {
                return fade.ThroughBlack?.Value == true
                    ? SlideTransition.FadeThroughBlack
                    : SlideTransition.Fade;
            }

            WipeTransition? wipe = transition.GetFirstChild<WipeTransition>();
            if (wipe != null) {
                TransitionSlideDirectionValues? direction = wipe.Direction?.Value;
                if (direction == TransitionSlideDirectionValues.Up) return SlideTransition.WipeUp;
                if (direction == TransitionSlideDirectionValues.Right) return SlideTransition.WipeRight;
                if (direction == TransitionSlideDirectionValues.Down) return SlideTransition.WipeDown;
                return SlideTransition.Wipe;
            }

            BlindsTransition? blinds = transition.GetFirstChild<BlindsTransition>();
            if (blinds != null) {
                return blinds.Direction?.Value == DirectionValues.Vertical
                    ? SlideTransition.BlindsVertical
                    : SlideTransition.BlindsHorizontal;
            }

            CheckerTransition? checker = transition.GetFirstChild<CheckerTransition>();
            if (checker != null) {
                return checker.Direction?.Value == DirectionValues.Vertical
                    ? SlideTransition.CheckerVertical
                    : SlideTransition.CheckerHorizontal;
            }

            CombTransition? comb = transition.GetFirstChild<CombTransition>();
            if (comb != null) {
                return comb.Direction?.Value == DirectionValues.Vertical
                    ? SlideTransition.CombVertical
                    : SlideTransition.CombHorizontal;
            }

            CoverTransition? cover = transition.GetFirstChild<CoverTransition>();
            if (cover != null) {
                return MapEightDirection(cover.Direction?.Value,
                    SlideTransition.CoverLeft, SlideTransition.CoverUp,
                    SlideTransition.CoverRight, SlideTransition.CoverDown,
                    SlideTransition.CoverLeftUp, SlideTransition.CoverRightUp,
                    SlideTransition.CoverLeftDown, SlideTransition.CoverRightDown);
            }

            PullTransition? pull = transition.GetFirstChild<PullTransition>();
            if (pull != null) {
                return MapEightDirection(pull.Direction?.Value,
                    SlideTransition.UncoverLeft, SlideTransition.UncoverUp,
                    SlideTransition.UncoverRight, SlideTransition.UncoverDown,
                    SlideTransition.UncoverLeftUp, SlideTransition.UncoverRightUp,
                    SlideTransition.UncoverLeftDown, SlideTransition.UncoverRightDown);
            }

            RandomBarTransition? randomBars =
                transition.GetFirstChild<RandomBarTransition>();
            if (randomBars != null) {
                return randomBars.Direction?.Value == DirectionValues.Vertical
                    ? SlideTransition.RandomBarsVertical
                    : SlideTransition.RandomBarsHorizontal;
            }

            StripsTransition? strips = transition.GetFirstChild<StripsTransition>();
            if (strips != null) {
                TransitionCornerDirectionValues? direction = strips.Direction?.Value;
                if (direction == TransitionCornerDirectionValues.RightUp) return SlideTransition.StripsRightUp;
                if (direction == TransitionCornerDirectionValues.LeftDown) return SlideTransition.StripsLeftDown;
                if (direction == TransitionCornerDirectionValues.RightDown) return SlideTransition.StripsRightDown;
                return SlideTransition.StripsLeftUp;
            }

            PushTransition? push = transition.GetFirstChild<PushTransition>();
            if (push != null) {
                TransitionSlideDirectionValues? direction = push.Direction?.Value;
                if (direction == TransitionSlideDirectionValues.Up) return SlideTransition.PushUp;
                if (direction == TransitionSlideDirectionValues.Down) return SlideTransition.PushDown;
                if (direction == TransitionSlideDirectionValues.Right) return SlideTransition.PushRight;
                return SlideTransition.PushLeft;
            }

            ZoomTransition? zoom = transition.GetFirstChild<ZoomTransition>();
            if (zoom != null) {
                return zoom.Direction?.Value == TransitionInOutDirectionValues.In
                    ? SlideTransition.BoxIn
                    : SlideTransition.BoxOut;
            }

            SplitTransition? split = transition.GetFirstChild<SplitTransition>();
            if (split != null) {
                bool vertical = split.Orientation?.Value == DirectionValues.Vertical;
                bool inward = split.Direction?.Value == TransitionInOutDirectionValues.In;
                return vertical
                    ? inward ? SlideTransition.SplitVerticalIn : SlideTransition.SplitVerticalOut
                    : inward ? SlideTransition.SplitHorizontalIn : SlideTransition.SplitHorizontalOut;
            }

            WheelTransition? wheel = transition.GetFirstChild<WheelTransition>();
            if (wheel != null) {
                return wheel.Spokes?.Value switch {
                    1 => SlideTransition.WheelOneSpoke,
                    2 => SlideTransition.WheelTwoSpokes,
                    3 => SlideTransition.WheelThreeSpokes,
                    4 => SlideTransition.WheelFourSpokes,
                    8 => SlideTransition.WheelEightSpokes,
                    _ => null
                };
            }

            if (transition.GetFirstChild<RandomTransition>() != null) {
                return SlideTransition.Random;
            }
            if (transition.GetFirstChild<DissolveTransition>() != null) {
                return SlideTransition.Dissolve;
            }
            if (transition.GetFirstChild<DiamondTransition>() != null) {
                return SlideTransition.Diamond;
            }
            if (transition.GetFirstChild<PlusTransition>() != null) {
                return SlideTransition.Plus;
            }
            if (transition.GetFirstChild<WedgeTransition>() != null) {
                return SlideTransition.Wedge;
            }
            if (transition.GetFirstChild<NewsflashTransition>() != null) {
                return SlideTransition.Newsflash;
            }
            if (transition.GetFirstChild<CircleTransition>() != null) {
                return SlideTransition.Circle;
            }

            CutTransition? cut = transition.GetFirstChild<CutTransition>();
            if (cut != null) {
                return cut.ThroughBlack?.Value == true
                    ? SlideTransition.CutThroughBlack
                    : SlideTransition.Cut;
            }

            return null;
        }

        private static SlideTransition MapEightDirection(string? direction,
            SlideTransition left, SlideTransition up, SlideTransition right,
            SlideTransition down, SlideTransition leftUp, SlideTransition rightUp,
            SlideTransition leftDown, SlideTransition rightDown) => direction switch {
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
            SlideTransition transition) => transition switch {
                SlideTransition.Fade => new FadeTransition { ThroughBlack = false },
                SlideTransition.FadeThroughBlack => new FadeTransition { ThroughBlack = true },
                SlideTransition.Wipe => CreateWipe(TransitionSlideDirectionValues.Left),
                SlideTransition.WipeUp => CreateWipe(TransitionSlideDirectionValues.Up),
                SlideTransition.WipeRight => CreateWipe(TransitionSlideDirectionValues.Right),
                SlideTransition.WipeDown => CreateWipe(TransitionSlideDirectionValues.Down),
                SlideTransition.BlindsVertical => new BlindsTransition { Direction = DirectionValues.Vertical },
                SlideTransition.BlindsHorizontal => new BlindsTransition { Direction = DirectionValues.Horizontal },
                SlideTransition.CheckerHorizontal => new CheckerTransition { Direction = DirectionValues.Horizontal },
                SlideTransition.CheckerVertical => new CheckerTransition { Direction = DirectionValues.Vertical },
                SlideTransition.CombHorizontal => new CombTransition { Direction = DirectionValues.Horizontal },
                SlideTransition.CombVertical => new CombTransition { Direction = DirectionValues.Vertical },
                SlideTransition.CoverLeft => CreateCover("l"),
                SlideTransition.CoverUp => CreateCover("u"),
                SlideTransition.CoverRight => CreateCover("r"),
                SlideTransition.CoverDown => CreateCover("d"),
                SlideTransition.CoverLeftUp => CreateCover("lu"),
                SlideTransition.CoverRightUp => CreateCover("ru"),
                SlideTransition.CoverLeftDown => CreateCover("ld"),
                SlideTransition.CoverRightDown => CreateCover("rd"),
                SlideTransition.UncoverLeft => CreatePull("l"),
                SlideTransition.UncoverUp => CreatePull("u"),
                SlideTransition.UncoverRight => CreatePull("r"),
                SlideTransition.UncoverDown => CreatePull("d"),
                SlideTransition.UncoverLeftUp => CreatePull("lu"),
                SlideTransition.UncoverRightUp => CreatePull("ru"),
                SlideTransition.UncoverLeftDown => CreatePull("ld"),
                SlideTransition.UncoverRightDown => CreatePull("rd"),
                SlideTransition.RandomBarsHorizontal => new RandomBarTransition { Direction = DirectionValues.Horizontal },
                SlideTransition.RandomBarsVertical => new RandomBarTransition { Direction = DirectionValues.Vertical },
                SlideTransition.StripsLeftUp => CreateStrips(TransitionCornerDirectionValues.LeftUp),
                SlideTransition.StripsRightUp => CreateStrips(TransitionCornerDirectionValues.RightUp),
                SlideTransition.StripsLeftDown => CreateStrips(TransitionCornerDirectionValues.LeftDown),
                SlideTransition.StripsRightDown => CreateStrips(TransitionCornerDirectionValues.RightDown),
                SlideTransition.PushLeft => CreatePush(TransitionSlideDirectionValues.Left),
                SlideTransition.PushUp => CreatePush(TransitionSlideDirectionValues.Up),
                SlideTransition.PushRight => CreatePush(TransitionSlideDirectionValues.Right),
                SlideTransition.PushDown => CreatePush(TransitionSlideDirectionValues.Down),
                SlideTransition.BoxOut => new ZoomTransition { Direction = TransitionInOutDirectionValues.Out },
                SlideTransition.BoxIn => new ZoomTransition { Direction = TransitionInOutDirectionValues.In },
                SlideTransition.SplitHorizontalOut => CreateSplit(DirectionValues.Horizontal, TransitionInOutDirectionValues.Out),
                SlideTransition.SplitHorizontalIn => CreateSplit(DirectionValues.Horizontal, TransitionInOutDirectionValues.In),
                SlideTransition.SplitVerticalOut => CreateSplit(DirectionValues.Vertical, TransitionInOutDirectionValues.Out),
                SlideTransition.SplitVerticalIn => CreateSplit(DirectionValues.Vertical, TransitionInOutDirectionValues.In),
                SlideTransition.Random => new RandomTransition(),
                SlideTransition.Dissolve => new DissolveTransition(),
                SlideTransition.Diamond => new DiamondTransition(),
                SlideTransition.Plus => new PlusTransition(),
                SlideTransition.Wedge => new WedgeTransition(),
                SlideTransition.Newsflash => new NewsflashTransition(),
                SlideTransition.WheelOneSpoke => new WheelTransition { Spokes = 1U },
                SlideTransition.WheelTwoSpokes => new WheelTransition { Spokes = 2U },
                SlideTransition.WheelThreeSpokes => new WheelTransition { Spokes = 3U },
                SlideTransition.WheelFourSpokes => new WheelTransition { Spokes = 4U },
                SlideTransition.WheelEightSpokes => new WheelTransition { Spokes = 8U },
                SlideTransition.Circle => new CircleTransition(),
                SlideTransition.Cut => new CutTransition { ThroughBlack = false },
                SlideTransition.CutThroughBlack => new CutTransition { ThroughBlack = true },
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
