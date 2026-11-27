using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        /// The transition branches of a slide element, as <see cref="GetTransitionBranches()"/> reads them: its
        /// <c>p:transition</c>, or the transition of each choice and of the fallback of the first alternate-content block
        /// that carries one, with none where a branch has no transition.
        /// </summary>
        internal static IReadOnlyList<OfficeXmlNode> GetTransitionBranches(OfficeXmlNode slide) {
            OfficeXmlNode transition = slide.Child(OfficeXmlNamespaces.PresentationML, "transition");
            if (!transition.IsNone) return new[] { transition };
            foreach (OfficeXmlNode content in slide.Children(OfficeXmlNamespaces.MarkupCompatibility, "AlternateContent")) {
                OfficeXmlNode fallback = content.Child(OfficeXmlNamespaces.MarkupCompatibility, "Fallback");
                bool carriesTransition = !fallback.Child(OfficeXmlNamespaces.PresentationML, "transition").IsNone;
                foreach (OfficeXmlNode choice in content.Children(OfficeXmlNamespaces.MarkupCompatibility, "Choice")) {
                    if (!choice.Child(OfficeXmlNamespaces.PresentationML, "transition").IsNone) carriesTransition = true;
                }
                if (!carriesTransition) continue;
                var branches = new List<OfficeXmlNode>();
                foreach (OfficeXmlNode choice in content.Children(OfficeXmlNamespaces.MarkupCompatibility, "Choice")) {
                    branches.Add(choice.Child(OfficeXmlNamespaces.PresentationML, "transition"));
                }
                if (!fallback.IsNone) branches.Add(fallback.Child(OfficeXmlNamespaces.PresentationML, "transition"));
                return branches;
            }
            return Array.Empty<OfficeXmlNode>();
        }

        /// <summary>The effect a <c>p:transition</c> element describes, read as the slide reads its own transition.</summary>
        internal static PowerPointSlideTransition GetTransitionValue(OfficeXmlNode transition) {
            if (transition.IsNone) throw new ArgumentNullException(nameof(transition));
            PowerPointSlideTransition? classic = GetClassicTransition(transition);
            if (classic.HasValue) return classic.Value;
            if (!transition.Child(P14Namespace, "flash").IsNone) return PowerPointSlideTransition.Flash;
            OfficeXmlNode warp = transition.Child(P14Namespace, "warp");
            if (!warp.IsNone) {
                return warp.Attribute("dir") == "out" ? PowerPointSlideTransition.WarpOut : PowerPointSlideTransition.WarpIn;
            }
            if (!transition.Child(P14Namespace, "prism").IsNone) return PowerPointSlideTransition.Prism;
            OfficeXmlNode ferris = transition.Child(P14Namespace, "ferris");
            if (!ferris.IsNone) {
                return ferris.Attribute("dir") == "r" ? PowerPointSlideTransition.FerrisRight : PowerPointSlideTransition.FerrisLeft;
            }
            foreach (OfficeXmlNode child in transition.Children()) {
                if (child.Is(P159Namespace, "morph")
                    || child.Is("http://schemas.microsoft.com/office/powerpoint/2012/main", "prstTrans")
                    && string.Equals(child.Attribute("prst"), "morph", StringComparison.OrdinalIgnoreCase)) {
                    return PowerPointSlideTransition.Morph;
                }
            }
            return PowerPointSlideTransition.None;
        }

        private static PowerPointSlideTransition? GetClassicTransition(OfficeXmlNode transition) {
            OfficeXmlNode Classic(string localName) => transition.Child(OfficeXmlNamespaces.PresentationML, localName);
            static bool Vertical(OfficeXmlNode element, string attribute = "dir") => element.Attribute(attribute) == "vert";
            static bool ThroughBlack(OfficeXmlNode element) => OfficeXmlViewValues.ParseBoolean(element.Attribute("thruBlk")) == true;

            OfficeXmlNode fade = Classic("fade");
            if (!fade.IsNone) return ThroughBlack(fade) ? PowerPointSlideTransition.FadeThroughBlack : PowerPointSlideTransition.Fade;

            OfficeXmlNode wipe = Classic("wipe");
            if (!wipe.IsNone) {
                return wipe.Attribute("dir") switch {
                    "u" => PowerPointSlideTransition.WipeUp,
                    "r" => PowerPointSlideTransition.WipeRight,
                    "d" => PowerPointSlideTransition.WipeDown,
                    _ => PowerPointSlideTransition.Wipe
                };
            }

            OfficeXmlNode blinds = Classic("blinds");
            if (!blinds.IsNone) return Vertical(blinds) ? PowerPointSlideTransition.BlindsVertical : PowerPointSlideTransition.BlindsHorizontal;
            OfficeXmlNode checker = Classic("checker");
            if (!checker.IsNone) return Vertical(checker) ? PowerPointSlideTransition.CheckerVertical : PowerPointSlideTransition.CheckerHorizontal;
            OfficeXmlNode comb = Classic("comb");
            if (!comb.IsNone) return Vertical(comb) ? PowerPointSlideTransition.CombVertical : PowerPointSlideTransition.CombHorizontal;

            OfficeXmlNode cover = Classic("cover");
            if (!cover.IsNone) {
                return MapEightDirection(cover.Attribute("dir"),
                    PowerPointSlideTransition.CoverLeft, PowerPointSlideTransition.CoverUp,
                    PowerPointSlideTransition.CoverRight, PowerPointSlideTransition.CoverDown,
                    PowerPointSlideTransition.CoverLeftUp, PowerPointSlideTransition.CoverRightUp,
                    PowerPointSlideTransition.CoverLeftDown, PowerPointSlideTransition.CoverRightDown);
            }
            OfficeXmlNode pull = Classic("pull");
            if (!pull.IsNone) {
                return MapEightDirection(pull.Attribute("dir"),
                    PowerPointSlideTransition.UncoverLeft, PowerPointSlideTransition.UncoverUp,
                    PowerPointSlideTransition.UncoverRight, PowerPointSlideTransition.UncoverDown,
                    PowerPointSlideTransition.UncoverLeftUp, PowerPointSlideTransition.UncoverRightUp,
                    PowerPointSlideTransition.UncoverLeftDown, PowerPointSlideTransition.UncoverRightDown);
            }

            OfficeXmlNode randomBars = Classic("randomBar");
            if (!randomBars.IsNone) return Vertical(randomBars) ? PowerPointSlideTransition.RandomBarsVertical : PowerPointSlideTransition.RandomBarsHorizontal;

            OfficeXmlNode strips = Classic("strips");
            if (!strips.IsNone) {
                return strips.Attribute("dir") switch {
                    "ru" => PowerPointSlideTransition.StripsRightUp,
                    "ld" => PowerPointSlideTransition.StripsLeftDown,
                    "rd" => PowerPointSlideTransition.StripsRightDown,
                    _ => PowerPointSlideTransition.StripsLeftUp
                };
            }

            OfficeXmlNode push = Classic("push");
            if (!push.IsNone) {
                return push.Attribute("dir") switch {
                    "u" => PowerPointSlideTransition.PushUp,
                    "d" => PowerPointSlideTransition.PushDown,
                    "r" => PowerPointSlideTransition.PushRight,
                    _ => PowerPointSlideTransition.PushLeft
                };
            }

            OfficeXmlNode zoom = Classic("zoom");
            if (!zoom.IsNone) return zoom.Attribute("dir") == "in" ? PowerPointSlideTransition.BoxIn : PowerPointSlideTransition.BoxOut;

            OfficeXmlNode split = Classic("split");
            if (!split.IsNone) {
                bool vertical = Vertical(split, "orient");
                bool inward = split.Attribute("dir") == "in";
                return vertical
                    ? inward ? PowerPointSlideTransition.SplitVerticalIn : PowerPointSlideTransition.SplitVerticalOut
                    : inward ? PowerPointSlideTransition.SplitHorizontalIn : PowerPointSlideTransition.SplitHorizontalOut;
            }

            OfficeXmlNode wheel = Classic("wheel");
            if (!wheel.IsNone) {
                return OfficeXmlViewValues.ParseUInt32(wheel.Attribute("spokes")) switch {
                    1U => PowerPointSlideTransition.WheelOneSpoke,
                    2U => PowerPointSlideTransition.WheelTwoSpokes,
                    3U => PowerPointSlideTransition.WheelThreeSpokes,
                    4U => PowerPointSlideTransition.WheelFourSpokes,
                    8U => PowerPointSlideTransition.WheelEightSpokes,
                    _ => null
                };
            }

            if (!Classic("random").IsNone) return PowerPointSlideTransition.Random;
            if (!Classic("dissolve").IsNone) return PowerPointSlideTransition.Dissolve;
            if (!Classic("diamond").IsNone) return PowerPointSlideTransition.Diamond;
            if (!Classic("plus").IsNone) return PowerPointSlideTransition.Plus;
            if (!Classic("wedge").IsNone) return PowerPointSlideTransition.Wedge;
            if (!Classic("newsflash").IsNone) return PowerPointSlideTransition.Newsflash;
            if (!Classic("circle").IsNone) return PowerPointSlideTransition.Circle;

            OfficeXmlNode cut = Classic("cut");
            if (!cut.IsNone) return ThroughBlack(cut) ? PowerPointSlideTransition.CutThroughBlack : PowerPointSlideTransition.Cut;
            return null;
        }
    }
}
