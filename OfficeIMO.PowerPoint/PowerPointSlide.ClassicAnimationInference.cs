using System.Globalization;
using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private static IReadOnlyList<PowerPointClassicAnimation> InferClassicAnimations(OfficeXmlNode timing) {
            if (timing.IsNone) return Array.Empty<PowerPointClassicAnimation>();
            OfficeXmlNode[] nodes = TimingNodesNamed(timing, null).ToArray();
            if (nodes.Any(node => node.LocalName is "anim" or "animClr" or "animMotion" or "animRot" or "animScale" or "cmd" or "audio" or "video")) {
                return Array.Empty<PowerPointClassicAnimation>();
            }

            const int p = ClassicPresentationNamespace;
            var builds = new Dictionary<uint, OfficeXmlNode>();
            foreach (OfficeXmlNode build in ClassicChildren(timing.Child(p, "bldLst"))) {
                if (!build.Is(p, "bldP") || !TryReadClassicShapeId(build, out uint shapeId) || builds.ContainsKey(shapeId)) {
                    return Array.Empty<PowerPointClassicAnimation>();
                }
                builds.Add(shapeId, build);
            }

            var result = new List<PowerPointClassicAnimation>();
            bool hasAfterEffect = nodes.Any(node => node.LocalName == "cTn" && ReadClassicTimingBoolean(node, "afterEffect"));
            foreach (OfficeXmlNode effect in nodes.Where(node => node.LocalName == "animEffect")) {
                OfficeXmlNode[] targets = TimingNodesNamed(effect, "spTgt").ToArray();
                OfficeXmlNode effectTime = TimingNodesNamed(effect, "cTn").FirstOrDefault();
                OfficeXmlNode ownerTime = effect.Parent.AncestorOrSelf(p, "cTn");
                string? nodeType = ReadClassicTimingToken(ownerTime, "nodeType");
                if (targets.Length != 1 || !TryReadClassicShapeId(targets[0], out uint shapeId)
                    || effectTime.IsNone || ownerTime.IsNone || hasAfterEffect || ReadClassicTimingToken(effect, "transition") == "out"
                    || nodeType is not ("clickEffect" or "afterEffect")
                    || ReadClassicTimingToken(effectTime, "presetClass") is not (null or "entr")
                    || ReadClassicTimingToken(ownerTime, "presetClass") is not (null or "entr")
                    || !TryMapClassicAnimationFilter(effect.Attribute("filter"), out PowerPointClassicAnimationEffect mappedEffect, out byte direction)) {
                    return Array.Empty<PowerPointClassicAnimation>();
                }

                OfficeXmlNode condition = ownerTime.Child(p, "stCondLst").Child(p, "cond");
                bool automatic = nodeType == "afterEffect";
                int delay = 0;
                string? delayValue = condition.Attribute("delay");
                string? trigger = ReadClassicTimingToken(condition, "evt");
                if (automatic) {
                    if (trigger == "onClick" || !string.IsNullOrEmpty(delayValue)
                        && !int.TryParse(delayValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out delay)) {
                        return Array.Empty<PowerPointClassicAnimation>();
                    }
                } else if (trigger is not (null or "onClick") || !string.IsNullOrEmpty(delayValue) && delayValue != "0") {
                    return Array.Empty<PowerPointClassicAnimation>();
                }

                PowerPointClassicAnimationBuildType buildType = PowerPointClassicAnimationBuildType.AsOneObject;
                bool reverse = ReadClassicTimingBoolean(effectTime, "autoRev") || ReadClassicTimingBoolean(ownerTime, "autoRev");
                bool animateBackground = false;
                if (builds.TryGetValue(shapeId, out OfficeXmlNode build)) {
                    animateBackground = ReadClassicTimingBoolean(build, "animBg");
                    reverse |= ReadClassicTimingBoolean(build, "rev");
                    string? buildKind = ReadClassicTimingToken(build, "build");
                    if (buildKind == "p") {
                        uint level = build.Attribute("bldLvl") is { } token ? XmlConvert.ToUInt32(token) : 1U;
                        if (level is < 1 or > 5) return Array.Empty<PowerPointClassicAnimation>();
                        buildType = (PowerPointClassicAnimationBuildType)checked((byte)(level + 1U));
                    } else if (buildKind is not (null or "whole" or "allAtOnce")) {
                        return Array.Empty<PowerPointClassicAnimation>();
                    }
                }

                result.Add(new PowerPointClassicAnimation(shapeId, mappedEffect, direction, buildType, automatic,
                    Math.Max(0, delay), result.Count, reverse, animateBackground,
                    PowerPointClassicAnimationAfterEffect.None, PowerPointClassicTextBuild.AllAtOnce, rawDimColor: 0));
            }
            return HasOnlyClassicStandardActions(nodes, result) ? result : Array.Empty<PowerPointClassicAnimation>();
        }

        private static bool TryReadClassicShapeId(OfficeXmlNode node, out uint shapeId) =>
            uint.TryParse(node.Attribute("spid"), NumberStyles.Integer, CultureInfo.InvariantCulture, out shapeId);

        private static bool ReadClassicTimingBoolean(OfficeXmlNode node, string attribute) =>
            node.Attribute(attribute) is { } value && XmlConvert.ToBoolean(value);

        private static IEnumerable<OfficeXmlNode> ClassicChildren(OfficeXmlNode node) {
            foreach (OfficeXmlNode child in node.ChildNodes()) {
                if (child.Tree is { } tree && tree.KindOf(child.Node) == OfficeXmlTreeNodeKind.Text
                    && !tree.IsCData(child.Node) && string.IsNullOrWhiteSpace(tree.GetText(child.Node))) continue;
                yield return child;
            }
        }

        private static string? ReadClassicTimingToken(OfficeXmlNode node, string attribute) {
            string? token = node.Attribute(attribute);
            if (token == null) return null;
            bool valid = attribute switch {
                "nodeType" => token is "clickEffect" or "withEffect" or "afterEffect" or "mainSeq" or "interactiveSeq" or "clickPar"
                    or "withGroup" or "afterGroup" or "tmRoot",
                "transition" => token is "in" or "out" or "none",
                "presetClass" => token is "entr" or "exit" or "emph" or "path" or "verb" or "mediacall",
                "evt" => token is "none" or "onBegin" or "onEnd" or "begin" or "end" or "onClick" or "onDblClick" or "onMouseOver"
                    or "onMouseOut" or "onNext" or "onPrev" or "onStopAudio" or "onMediaBookmark",
                "build" => token is "allAtOnce" or "p" or "cust" or "whole",
                "type" => token is "evt" or "call" or "verb",
                _ => false
            };
            return valid ? token : throw new FormatException("The text value is not a valid enumeration value.");
        }

        private static bool HasOnlyClassicStandardActions(IReadOnlyList<OfficeXmlNode> timing, IReadOnlyList<PowerPointClassicAnimation> animations) {
            OfficeXmlNode[] sets = timing.Where(node => node.LocalName == "set").ToArray();
            OfficeXmlNode[] colors = timing.Where(node => node.LocalName == "animClr").ToArray();
            OfficeXmlNode[] sounds = timing.Where(node => node.LocalName == "audio").ToArray();
            OfficeXmlNode[] commands = timing.Where(node => node.LocalName == "cmd").ToArray();
            const int p = ClassicPresentationNamespace;
            foreach (OfficeXmlNode set in sets) {
                OfficeXmlNode behavior = set.Child(p, "cBhvr");
                OfficeXmlNode target = behavior.Child(p, "tgtEl").Child(p, "spTgt");
                OfficeXmlNode to = set.Child(p, "to");
                OfficeXmlNode value = to.Child(p, "strVal");
                if (!TryReadClassicShapeId(target, out uint shapeId)
                    || animations.FirstOrDefault(animation => animation.ShapeId == shapeId) is not PowerPointClassicAnimation animation
                    || ClassicChildren(set).Count() != 2
                    || !HasClassicBehaviorAttribute(behavior, "style.visibility")
                    || value.IsNone || ClassicChildren(to).Count() != 1
                    || !string.Equals(value.Attribute("val"), "visible", StringComparison.OrdinalIgnoreCase)
                    && (!string.Equals(value.Attribute("val"), "hidden", StringComparison.OrdinalIgnoreCase)
                        || animation.AfterEffect is not (PowerPointClassicAnimationAfterEffect.HideImmediately or PowerPointClassicAnimationAfterEffect.HideOnNextClick))
                    || !HasClassicEffectFor(FindClassicEffectOwner(set), shapeId)) {
                    return false;
                }
            }
            if (colors.Any(color => !animations.Any(animation => IsClassicDimAction(color, animation)))
                || sounds.Any(sound => !animations.Any(animation => IsClassicSoundAction(sound, animation)))
                || commands.Any(command => !animations.Any(animation => IsClassicStopSoundAction(command, animation)))) {
                return false;
            }
            foreach (PowerPointClassicAnimation animation in animations) {
                bool expectsHidden = animation.AfterEffect is PowerPointClassicAnimationAfterEffect.HideImmediately
                    or PowerPointClassicAnimationAfterEffect.HideOnNextClick;
                if (sets.Count(set => IsClassicHiddenAction(set, animation)) != (expectsHidden ? 1 : 0)
                    || colors.Count(color => IsClassicDimAction(color, animation)) != (animation.AfterEffect == PowerPointClassicAnimationAfterEffect.Dim ? 1 : 0)
                    || sounds.Count(sound => IsClassicSoundAction(sound, animation)) != (animation.PlaysSound && !string.IsNullOrEmpty(animation.SoundRelationshipId) ? 1 : 0)
                    || commands.Count(command => IsClassicStopSoundAction(command, animation)) != (animation.StopsSound ? 1 : 0)) {
                    return false;
                }
            }
            return true;
        }

        private static bool IsClassicHiddenAction(OfficeXmlNode set, PowerPointClassicAnimation animation) =>
            IsClassicShapeBehavior(set.Child(ClassicPresentationNamespace, "cBhvr"), animation.ShapeId, "style.visibility")
            && string.Equals(set.Child(ClassicPresentationNamespace, "to").Child(ClassicPresentationNamespace, "strVal").Attribute("val"),
                "hidden", StringComparison.OrdinalIgnoreCase);

        private static bool IsClassicDimAction(OfficeXmlNode color, PowerPointClassicAnimation animation) =>
            animation.AfterEffect == PowerPointClassicAnimationAfterEffect.Dim
            && IsClassicShapeBehavior(color.Child(ClassicPresentationNamespace, "cBhvr"), animation.ShapeId, "style.color");

        private static bool IsClassicSoundAction(OfficeXmlNode audio, PowerPointClassicAnimation animation) {
            const int p = ClassicPresentationNamespace;
            OfficeXmlNode target = audio.Child(p, "cMediaNode").Child(p, "tgtEl").Child(p, "sndTgt");
            return animation.PlaysSound && !string.IsNullOrEmpty(animation.SoundRelationshipId)
                && string.Equals(target.Attribute(OfficeXmlNamespaces.Relationships, "embed"),
                    animation.SoundRelationshipId, StringComparison.Ordinal);
        }

        private static bool IsClassicStopSoundAction(OfficeXmlNode command, PowerPointClassicAnimation animation) {
            OfficeXmlNode behavior = command.Child(ClassicPresentationNamespace, "cBhvr");
            return animation.StopsSound && ReadClassicTimingToken(command, "type") == "evt"
                && string.Equals(command.Attribute("cmd"), "onstopaudio", StringComparison.OrdinalIgnoreCase)
                && ClassicChildren(command).Count() == 1 && behavior.Exists
                && TimingNodesNamed(behavior, "sldTgt").Count() == 1 && !TimingNodesNamed(behavior, "spTgt").Any()
                && HasClassicEffectFor(FindClassicEffectOwner(command), animation.ShapeId);
        }

        private static bool IsClassicShapeBehavior(OfficeXmlNode behavior, uint shapeId, string attributeName) {
            const int p = ClassicPresentationNamespace;
            OfficeXmlNode target = behavior.Child(p, "tgtEl").Child(p, "spTgt");
            return string.Equals(target.Attribute("spid"), shapeId.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)
                && HasClassicBehaviorAttribute(behavior, attributeName);
        }

        private static bool HasClassicBehaviorAttribute(OfficeXmlNode behavior, string attributeName) {
            const int p = ClassicPresentationNamespace;
            OfficeXmlNode.ChildList attributes = behavior.Child(p, "attrNameLst").Children(p, "attrName");
            foreach (OfficeXmlNode attribute in attributes) {
                return attributes.Count == 1 && TimingNodesNamed(behavior, "spTgt").Count() == 1
                    && string.Equals(attribute.InnerText, attributeName, StringComparison.Ordinal);
            }
            return false;
        }

        private static OfficeXmlNode FindClassicEffectOwner(OfficeXmlNode action) {
            for (OfficeXmlNode node = action.Parent; node.Exists; node = node.Parent) {
                if (node.Is(ClassicPresentationNamespace, "cTn")
                    && ReadClassicTimingToken(node, "nodeType") is ("clickEffect" or "afterEffect")) return node;
            }
            return default;
        }

        private static bool HasClassicEffectFor(OfficeXmlNode owner, uint shapeId) => owner.Exists
            && TimingNodesNamed(owner, "animEffect").Count(effect =>
                TryReadClassicShapeId(TimingNodesNamed(effect, "spTgt").FirstOrDefault(), out uint target) && target == shapeId) == 1;

        private static bool TryMapClassicAnimationFilter(string? filter,
            out PowerPointClassicAnimationEffect effect, out byte direction) {
            effect = PowerPointClassicAnimationEffect.Cut;
            direction = 0;
            if (string.IsNullOrWhiteSpace(filter)) return false;
            string value = filter!.Replace(" ", string.Empty)
                .ToLowerInvariant();
            string name = value;
            string argument = string.Empty;
            int open = value.IndexOf('(');
            if (open >= 0 && value.EndsWith(")", StringComparison.Ordinal)) {
                name = value.Substring(0, open);
                argument = value.Substring(open + 1,
                    value.Length - open - 2);
            }
            switch (name) {
                case "cut":
                    effect = PowerPointClassicAnimationEffect.Cut;
                    direction = argument == "throughblack" ? (byte)1 : (byte)0;
                    return argument.Length == 0 || argument == "throughblack";
                case "random": effect = PowerPointClassicAnimationEffect.Random; return true;
                case "blinds":
                    effect = PowerPointClassicAnimationEffect.Blinds;
                    return TryMapOrientation(argument, verticalValue: 0,
                        out direction);
                case "checker":
                case "checkerboard":
                    effect = PowerPointClassicAnimationEffect.Checker;
                    if (argument is "across" or "horizontal") {
                        direction = 0;
                        return true;
                    }
                    if (argument is "down" or "vertical") {
                        direction = 1;
                        return true;
                    }
                    return false;
                case "cover":
                    effect = PowerPointClassicAnimationEffect.Cover;
                    return TryMapDirection(argument, true, out direction);
                case "dissolve":
                    effect = PowerPointClassicAnimationEffect.Dissolve;
                    return argument.Length == 0;
                case "fade":
                    effect = PowerPointClassicAnimationEffect.Fade;
                    return argument.Length == 0;
                case "pull":
                case "uncover":
                    effect = PowerPointClassicAnimationEffect.Pull;
                    return TryMapDirection(argument, true, out direction);
                case "randombar":
                case "randombars":
                    effect = PowerPointClassicAnimationEffect.RandomBars;
                    return TryMapOrientation(argument, verticalValue: 1,
                        out direction);
                case "strips":
                    effect = PowerPointClassicAnimationEffect.Strips;
                    return TryMapStripDirection(argument, out direction);
                case "wipe":
                    effect = PowerPointClassicAnimationEffect.Wipe;
                    return TryMapDirection(argument, false, out direction);
                case "box":
                case "zoom":
                    effect = PowerPointClassicAnimationEffect.Zoom;
                    if (argument == "out") { direction = 0; return true; }
                    if (argument == "in") { direction = 1; return true; }
                    return false;
                case "fly":
                    effect = PowerPointClassicAnimationEffect.Fly;
                    return TryMapFlyDirection(argument, out direction);
                case "split":
                    effect = PowerPointClassicAnimationEffect.Split;
                    return TryMapSplitDirection(argument, out direction);
                case "flash":
                    effect = PowerPointClassicAnimationEffect.Flash;
                    if (argument.Length == 0) return true;
                    return byte.TryParse(argument, NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out direction)
                        && direction <= 2;
                case "diamond":
                    effect = PowerPointClassicAnimationEffect.Diamond;
                    return argument.Length == 0 || argument == "in";
                case "plus":
                    effect = PowerPointClassicAnimationEffect.Plus;
                    return argument.Length == 0 || argument == "in";
                case "wedge":
                    effect = PowerPointClassicAnimationEffect.Wedge;
                    return argument.Length == 0;
                case "wheel":
                    effect = PowerPointClassicAnimationEffect.Wheel;
                    return byte.TryParse(argument, NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out direction)
                        && direction is 1 or 2 or 3 or 4 or 8;
                case "circle":
                    effect = PowerPointClassicAnimationEffect.Circle;
                    return argument.Length == 0 || argument == "in";
                default: return false;
            }
        }

        private static bool TryMapOrientation(string value, byte verticalValue,
            out byte direction) {
            direction = 0;
            if (value == "vertical") { direction = verticalValue; return true; }
            if (value == "horizontal") {
                direction = (byte)(1 - verticalValue);
                return true;
            }
            return false;
        }

        private static bool TryMapDirection(string value, bool allowCorners,
            out byte direction) {
            direction = value switch {
                "l" or "left" => 0, "u" or "up" => 1,
                "r" or "right" => 2, "d" or "down" => 3,
                "lu" or "leftup" or "upleft" => 4,
                "ru" or "rightup" or "upright" => 5,
                "ld" or "leftdown" or "downleft" => 6,
                "rd" or "rightdown" or "downright" => 7,
                _ => byte.MaxValue
            };
            return direction != byte.MaxValue
                && (allowCorners || direction <= 3);
        }

        private static bool TryMapStripDirection(string value,
            out byte direction) {
            if (!TryMapDirection(value, true, out direction)) return false;
            return direction >= 4;
        }

        private static bool TryMapFlyDirection(string value,
            out byte direction) {
            if (byte.TryParse(value, NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out direction)) {
                return direction <= 0x1C;
            }
            return TryMapDirection(value, true, out direction);
        }

        private static bool TryMapSplitDirection(string value,
            out byte direction) {
            direction = value switch {
                "horizontalout" => 0, "horizontalin" => 1,
                "verticalout" => 2, "verticalin" => 3,
                _ => byte.MaxValue
            };
            if (direction != byte.MaxValue) return true;
            return byte.TryParse(value, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out direction)
                && direction <= 3;
        }
    }
}
