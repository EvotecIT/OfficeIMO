using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private static PowerPointSlideTransition? GetClassicTransition(OfficeXmlNode transition) {
            OfficeXmlNode fade = transition.Child(TransitionPresentationNamespace, "fade");
            if (fade.Exists) {
                return fade.BooleanAttribute("thruBlk") == true
                    ? PowerPointSlideTransition.FadeThroughBlack
                    : PowerPointSlideTransition.Fade;
            }

            OfficeXmlNode wipe = transition.Child(TransitionPresentationNamespace, "wipe");
            if (wipe.Exists) {
                string? direction = wipe.Attribute("dir");
                if (direction == "u") return PowerPointSlideTransition.WipeUp;
                if (direction == "r") return PowerPointSlideTransition.WipeRight;
                if (direction == "d") return PowerPointSlideTransition.WipeDown;
                return PowerPointSlideTransition.Wipe;
            }

            OfficeXmlNode blinds = transition.Child(TransitionPresentationNamespace, "blinds");
            if (blinds.Exists) {
                return blinds.Attribute("dir") == "vert"
                    ? PowerPointSlideTransition.BlindsVertical
                    : PowerPointSlideTransition.BlindsHorizontal;
            }

            OfficeXmlNode checker = transition.Child(TransitionPresentationNamespace, "checker");
            if (checker.Exists) {
                return checker.Attribute("dir") == "vert"
                    ? PowerPointSlideTransition.CheckerVertical
                    : PowerPointSlideTransition.CheckerHorizontal;
            }

            OfficeXmlNode comb = transition.Child(TransitionPresentationNamespace, "comb");
            if (comb.Exists) {
                return comb.Attribute("dir") == "vert"
                    ? PowerPointSlideTransition.CombVertical
                    : PowerPointSlideTransition.CombHorizontal;
            }

            OfficeXmlNode cover = transition.Child(TransitionPresentationNamespace, "cover");
            if (cover.Exists) {
                return MapEightDirection(cover.Attribute("dir"),
                    PowerPointSlideTransition.CoverLeft, PowerPointSlideTransition.CoverUp,
                    PowerPointSlideTransition.CoverRight, PowerPointSlideTransition.CoverDown,
                    PowerPointSlideTransition.CoverLeftUp, PowerPointSlideTransition.CoverRightUp,
                    PowerPointSlideTransition.CoverLeftDown, PowerPointSlideTransition.CoverRightDown);
            }

            OfficeXmlNode pull = transition.Child(TransitionPresentationNamespace, "pull");
            if (pull.Exists) {
                return MapEightDirection(pull.Attribute("dir"),
                    PowerPointSlideTransition.UncoverLeft, PowerPointSlideTransition.UncoverUp,
                    PowerPointSlideTransition.UncoverRight, PowerPointSlideTransition.UncoverDown,
                    PowerPointSlideTransition.UncoverLeftUp, PowerPointSlideTransition.UncoverRightUp,
                    PowerPointSlideTransition.UncoverLeftDown, PowerPointSlideTransition.UncoverRightDown);
            }

            OfficeXmlNode randomBars =
                transition.Child(TransitionPresentationNamespace, "randomBar");
            if (randomBars.Exists) {
                return randomBars.Attribute("dir") == "vert"
                    ? PowerPointSlideTransition.RandomBarsVertical
                    : PowerPointSlideTransition.RandomBarsHorizontal;
            }

            OfficeXmlNode strips = transition.Child(TransitionPresentationNamespace, "strips");
            if (strips.Exists) {
                string? direction = strips.Attribute("dir");
                if (direction == "ru") return PowerPointSlideTransition.StripsRightUp;
                if (direction == "ld") return PowerPointSlideTransition.StripsLeftDown;
                if (direction == "rd") return PowerPointSlideTransition.StripsRightDown;
                return PowerPointSlideTransition.StripsLeftUp;
            }

            OfficeXmlNode push = transition.Child(TransitionPresentationNamespace, "push");
            if (push.Exists) {
                string? direction = push.Attribute("dir");
                if (direction == "u") return PowerPointSlideTransition.PushUp;
                if (direction == "d") return PowerPointSlideTransition.PushDown;
                if (direction == "r") return PowerPointSlideTransition.PushRight;
                return PowerPointSlideTransition.PushLeft;
            }

            OfficeXmlNode zoom = transition.Child(TransitionPresentationNamespace, "zoom");
            if (zoom.Exists) {
                return zoom.Attribute("dir") == "in"
                    ? PowerPointSlideTransition.BoxIn
                    : PowerPointSlideTransition.BoxOut;
            }

            OfficeXmlNode split = transition.Child(TransitionPresentationNamespace, "split");
            if (split.Exists) {
                bool vertical = split.Attribute("orient") == "vert";
                bool inward = split.Attribute("dir") == "in";
                return vertical
                    ? inward ? PowerPointSlideTransition.SplitVerticalIn : PowerPointSlideTransition.SplitVerticalOut
                    : inward ? PowerPointSlideTransition.SplitHorizontalIn : PowerPointSlideTransition.SplitHorizontalOut;
            }

            OfficeXmlNode wheel = transition.Child(TransitionPresentationNamespace, "wheel");
            if (wheel.Exists) {
                return wheel.UInt32Attribute("spokes") switch {
                    1 => PowerPointSlideTransition.WheelOneSpoke,
                    2 => PowerPointSlideTransition.WheelTwoSpokes,
                    3 => PowerPointSlideTransition.WheelThreeSpokes,
                    4 => PowerPointSlideTransition.WheelFourSpokes,
                    8 => PowerPointSlideTransition.WheelEightSpokes,
                    _ => null
                };
            }

            if (transition.Child(TransitionPresentationNamespace, "random").Exists) {
                return PowerPointSlideTransition.Random;
            }
            if (transition.Child(TransitionPresentationNamespace, "dissolve").Exists) {
                return PowerPointSlideTransition.Dissolve;
            }
            if (transition.Child(TransitionPresentationNamespace, "diamond").Exists) {
                return PowerPointSlideTransition.Diamond;
            }
            if (transition.Child(TransitionPresentationNamespace, "plus").Exists) {
                return PowerPointSlideTransition.Plus;
            }
            if (transition.Child(TransitionPresentationNamespace, "wedge").Exists) {
                return PowerPointSlideTransition.Wedge;
            }
            if (transition.Child(TransitionPresentationNamespace, "newsflash").Exists) {
                return PowerPointSlideTransition.Newsflash;
            }
            if (transition.Child(TransitionPresentationNamespace, "circle").Exists) {
                return PowerPointSlideTransition.Circle;
            }

            OfficeXmlNode cut = transition.Child(TransitionPresentationNamespace, "cut");
            if (cut.Exists) {
                return cut.BooleanAttribute("thruBlk") == true
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

        private static OfficeXmlNode CreateClassicTransition(OfficeXmlNode parent,
            PowerPointSlideTransition transition) => transition switch {
                PowerPointSlideTransition.Fade => CreateClassicTransitionPattern(parent, "fade", "thruBlk", "0"),
                PowerPointSlideTransition.FadeThroughBlack => CreateClassicTransitionPattern(parent, "fade", "thruBlk", "1"),
                PowerPointSlideTransition.Wipe => CreateClassicTransitionPattern(parent, "wipe", "dir", "l"),
                PowerPointSlideTransition.WipeUp => CreateClassicTransitionPattern(parent, "wipe", "dir", "u"),
                PowerPointSlideTransition.WipeRight => CreateClassicTransitionPattern(parent, "wipe", "dir", "r"),
                PowerPointSlideTransition.WipeDown => CreateClassicTransitionPattern(parent, "wipe", "dir", "d"),
                PowerPointSlideTransition.BlindsVertical => CreateClassicTransitionPattern(parent, "blinds", "dir", "vert"),
                PowerPointSlideTransition.BlindsHorizontal => CreateClassicTransitionPattern(parent, "blinds", "dir", "horz"),
                PowerPointSlideTransition.CheckerHorizontal => CreateClassicTransitionPattern(parent, "checker", "dir", "horz"),
                PowerPointSlideTransition.CheckerVertical => CreateClassicTransitionPattern(parent, "checker", "dir", "vert"),
                PowerPointSlideTransition.CombHorizontal => CreateClassicTransitionPattern(parent, "comb", "dir", "horz"),
                PowerPointSlideTransition.CombVertical => CreateClassicTransitionPattern(parent, "comb", "dir", "vert"),
                PowerPointSlideTransition.CoverLeft => CreateClassicTransitionPattern(parent, "cover", "dir", "l"),
                PowerPointSlideTransition.CoverUp => CreateClassicTransitionPattern(parent, "cover", "dir", "u"),
                PowerPointSlideTransition.CoverRight => CreateClassicTransitionPattern(parent, "cover", "dir", "r"),
                PowerPointSlideTransition.CoverDown => CreateClassicTransitionPattern(parent, "cover", "dir", "d"),
                PowerPointSlideTransition.CoverLeftUp => CreateClassicTransitionPattern(parent, "cover", "dir", "lu"),
                PowerPointSlideTransition.CoverRightUp => CreateClassicTransitionPattern(parent, "cover", "dir", "ru"),
                PowerPointSlideTransition.CoverLeftDown => CreateClassicTransitionPattern(parent, "cover", "dir", "ld"),
                PowerPointSlideTransition.CoverRightDown => CreateClassicTransitionPattern(parent, "cover", "dir", "rd"),
                PowerPointSlideTransition.UncoverLeft => CreateClassicTransitionPattern(parent, "pull", "dir", "l"),
                PowerPointSlideTransition.UncoverUp => CreateClassicTransitionPattern(parent, "pull", "dir", "u"),
                PowerPointSlideTransition.UncoverRight => CreateClassicTransitionPattern(parent, "pull", "dir", "r"),
                PowerPointSlideTransition.UncoverDown => CreateClassicTransitionPattern(parent, "pull", "dir", "d"),
                PowerPointSlideTransition.UncoverLeftUp => CreateClassicTransitionPattern(parent, "pull", "dir", "lu"),
                PowerPointSlideTransition.UncoverRightUp => CreateClassicTransitionPattern(parent, "pull", "dir", "ru"),
                PowerPointSlideTransition.UncoverLeftDown => CreateClassicTransitionPattern(parent, "pull", "dir", "ld"),
                PowerPointSlideTransition.UncoverRightDown => CreateClassicTransitionPattern(parent, "pull", "dir", "rd"),
                PowerPointSlideTransition.RandomBarsHorizontal => CreateClassicTransitionPattern(parent, "randomBar", "dir", "horz"),
                PowerPointSlideTransition.RandomBarsVertical => CreateClassicTransitionPattern(parent, "randomBar", "dir", "vert"),
                PowerPointSlideTransition.StripsLeftUp => CreateClassicTransitionPattern(parent, "strips", "dir", "lu"),
                PowerPointSlideTransition.StripsRightUp => CreateClassicTransitionPattern(parent, "strips", "dir", "ru"),
                PowerPointSlideTransition.StripsLeftDown => CreateClassicTransitionPattern(parent, "strips", "dir", "ld"),
                PowerPointSlideTransition.StripsRightDown => CreateClassicTransitionPattern(parent, "strips", "dir", "rd"),
                PowerPointSlideTransition.PushLeft => CreateClassicTransitionPattern(parent, "push", "dir", "l"),
                PowerPointSlideTransition.PushUp => CreateClassicTransitionPattern(parent, "push", "dir", "u"),
                PowerPointSlideTransition.PushRight => CreateClassicTransitionPattern(parent, "push", "dir", "r"),
                PowerPointSlideTransition.PushDown => CreateClassicTransitionPattern(parent, "push", "dir", "d"),
                PowerPointSlideTransition.BoxOut => CreateClassicTransitionPattern(parent, "zoom", "dir", "out"),
                PowerPointSlideTransition.BoxIn => CreateClassicTransitionPattern(parent, "zoom", "dir", "in"),
                PowerPointSlideTransition.SplitHorizontalOut => CreateClassicSplit(parent, "horz", "out"),
                PowerPointSlideTransition.SplitHorizontalIn => CreateClassicSplit(parent, "horz", "in"),
                PowerPointSlideTransition.SplitVerticalOut => CreateClassicSplit(parent, "vert", "out"),
                PowerPointSlideTransition.SplitVerticalIn => CreateClassicSplit(parent, "vert", "in"),
                PowerPointSlideTransition.Random => CreateClassicTransitionPattern(parent, "random"),
                PowerPointSlideTransition.Dissolve => CreateClassicTransitionPattern(parent, "dissolve"),
                PowerPointSlideTransition.Diamond => CreateClassicTransitionPattern(parent, "diamond"),
                PowerPointSlideTransition.Plus => CreateClassicTransitionPattern(parent, "plus"),
                PowerPointSlideTransition.Wedge => CreateClassicTransitionPattern(parent, "wedge"),
                PowerPointSlideTransition.Newsflash => CreateClassicTransitionPattern(parent, "newsflash"),
                PowerPointSlideTransition.WheelOneSpoke => CreateClassicTransitionPattern(parent, "wheel", "spokes", "1"),
                PowerPointSlideTransition.WheelTwoSpokes => CreateClassicTransitionPattern(parent, "wheel", "spokes", "2"),
                PowerPointSlideTransition.WheelThreeSpokes => CreateClassicTransitionPattern(parent, "wheel", "spokes", "3"),
                PowerPointSlideTransition.WheelFourSpokes => CreateClassicTransitionPattern(parent, "wheel", "spokes", "4"),
                PowerPointSlideTransition.WheelEightSpokes => CreateClassicTransitionPattern(parent, "wheel", "spokes", "8"),
                PowerPointSlideTransition.Circle => CreateClassicTransitionPattern(parent, "circle"),
                PowerPointSlideTransition.Cut => CreateClassicTransitionPattern(parent, "cut", "thruBlk", "0"),
                PowerPointSlideTransition.CutThroughBlack => CreateClassicTransitionPattern(parent, "cut", "thruBlk", "1"),
                _ => default
            };

        private static OfficeXmlNode CreateClassicTransitionPattern(OfficeXmlNode parent, string name,
            string? attributeName = null, string? attributeValue = null) {
            OfficeXmlNode transition = parent.CreateElement(TransitionPresentationNamespace, name);
            if (attributeName != null) transition.SetAttribute(attributeName, attributeValue);
            return transition;
        }

        private static OfficeXmlNode CreateClassicSplit(OfficeXmlNode parent, string orientation, string direction) {
            OfficeXmlNode transition = CreateClassicTransitionPattern(parent, "split", "orient", orientation);
            transition.SetAttribute("dir", direction);
            return transition;
        }
    }
}
