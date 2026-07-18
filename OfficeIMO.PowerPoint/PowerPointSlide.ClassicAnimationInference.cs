using System.Globalization;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private IReadOnlyList<PowerPointClassicAnimation> InferClassicAnimations() {
            Timing? timing = SlideRoot.Timing;
            if (timing == null) return Array.Empty<PowerPointClassicAnimation>();
            if (timing.Descendants().Any(element => element is Animate
                    or AnimateColor or AnimateMotion or AnimateRotation
                    or AnimateScale or Command or Audio or Video)) {
                return Array.Empty<PowerPointClassicAnimation>();
            }
            BuildList? buildList = timing.GetFirstChild<BuildList>();
            if (buildList?.ChildElements.Any(element =>
                    element is not BuildParagraph) == true) {
                return Array.Empty<PowerPointClassicAnimation>();
            }
            var builds = new Dictionary<uint, BuildParagraph>();
            foreach (BuildParagraph build in buildList?.Elements<BuildParagraph>()
                         ?? Enumerable.Empty<BuildParagraph>()) {
                if (!uint.TryParse(build.ShapeId?.Value, NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out uint shapeId)
                    || builds.ContainsKey(shapeId)) {
                    return Array.Empty<PowerPointClassicAnimation>();
                }
                builds.Add(shapeId, build);
            }

            var result = new List<PowerPointClassicAnimation>();
            bool hasAfterEffect = timing.Descendants<CommonTimeNode>().Any(node =>
                node.AfterEffect?.Value == true);
            foreach (AnimateEffect effect in timing.Descendants<AnimateEffect>()) {
                ShapeTarget? target = effect.Descendants<ShapeTarget>().FirstOrDefault();
                CommonTimeNode? effectTime = effect.Descendants<CommonTimeNode>()
                    .FirstOrDefault();
                CommonTimeNode? ownerTime = effect.Ancestors<CommonTimeNode>()
                    .FirstOrDefault();
                TimeNodeValues? nodeType = ownerTime?.NodeType?.Value;
                if (!uint.TryParse(target?.ShapeId?.Value, NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out uint shapeId)
                    || effectTime == null || ownerTime == null
                    || hasAfterEffect
                    || effect.Transition?.Value == AnimateEffectTransitionValues.Out
                    || nodeType != TimeNodeValues.ClickEffect
                    && nodeType != TimeNodeValues.AfterEffect
                    || effectTime.PresetClass?.Value is not null
                    && effectTime.PresetClass.Value !=
                        TimeNodePresetClassValues.Entrance
                    || ownerTime.PresetClass?.Value is not null
                    && ownerTime.PresetClass.Value !=
                        TimeNodePresetClassValues.Entrance
                    || effect.Descendants<ShapeTarget>().Count() != 1
                    || !TryMapClassicAnimationFilter(effect.Filter?.Value,
                        out PowerPointClassicAnimationEffect mappedEffect,
                        out byte direction)) {
                    return Array.Empty<PowerPointClassicAnimation>();
                }
                Condition? condition = ownerTime.GetFirstChild<StartConditionList>()?
                    .Elements<Condition>().FirstOrDefault();
                bool automatic = nodeType == TimeNodeValues.AfterEffect;
                int delay = 0;
                string? delayValue = condition?.Delay?.Value;
                if (automatic) {
                    if (condition?.Event?.Value == TriggerEventValues.OnClick
                        || !string.IsNullOrEmpty(delayValue)
                        && !int.TryParse(delayValue, NumberStyles.Integer,
                            CultureInfo.InvariantCulture, out delay)) {
                        return Array.Empty<PowerPointClassicAnimation>();
                    }
                } else if (condition?.Event?.Value is not null
                           && condition.Event.Value != TriggerEventValues.OnClick
                           || !string.IsNullOrEmpty(delayValue)
                           && delayValue != "0") {
                    return Array.Empty<PowerPointClassicAnimation>();
                }
                PowerPointClassicAnimationBuildType buildType =
                    PowerPointClassicAnimationBuildType.AsOneObject;
                bool reverse = effectTime.AutoReverse?.Value == true
                    || ownerTime.AutoReverse?.Value == true;
                bool animateBackground = false;
                if (builds.TryGetValue(shapeId, out BuildParagraph? build)) {
                    animateBackground = build.AnimateBackground?.Value == true;
                    reverse |= build.Reverse?.Value == true;
                    if (build.Build?.Value == ParagraphBuildValues.Paragraph) {
                        uint level = build.BuildLevel?.Value ?? 1U;
                        if (level is < 1 or > 5) {
                            return Array.Empty<PowerPointClassicAnimation>();
                        }
                        buildType = (PowerPointClassicAnimationBuildType)checked(
                            (byte)(level + 1U));
                    } else if (build.Build == null
                               || build.Build.Value == ParagraphBuildValues.Whole
                               || build.Build.Value == ParagraphBuildValues.AllAtOnce) {
                        buildType = PowerPointClassicAnimationBuildType.AsOneObject;
                    } else {
                        return Array.Empty<PowerPointClassicAnimation>();
                    }
                }
                result.Add(new PowerPointClassicAnimation(shapeId,
                    mappedEffect, direction, buildType, automatic,
                    Math.Max(0, delay), result.Count, reverse,
                    animateBackground,
                    PowerPointClassicAnimationAfterEffect.None,
                    PowerPointClassicTextBuild.AllAtOnce, rawDimColor: 0));
            }
            return HasOnlyClassicStandardActions(timing, result)
                ? result : Array.Empty<PowerPointClassicAnimation>();
        }

        private static bool HasOnlyClassicStandardActions(Timing timing,
            IReadOnlyList<PowerPointClassicAnimation> animations) {
            foreach (SetBehavior set in timing.Descendants<SetBehavior>()) {
                CommonBehavior? behavior = set.CommonBehavior;
                ShapeTarget? target = behavior?.GetFirstChild<TargetElement>()?
                    .GetFirstChild<ShapeTarget>();
                AttributeName[] attributes = behavior?
                    .GetFirstChild<AttributeNameList>()?
                    .Elements<AttributeName>().ToArray()
                    ?? Array.Empty<AttributeName>();
                StringVariantValue? value = set.ToVariantValue?
                    .GetFirstChild<StringVariantValue>();
                CommonTimeNode? owner = set.Ancestors<CommonTimeNode>()
                    .FirstOrDefault(node => node.NodeType?.Value ==
                            TimeNodeValues.ClickEffect
                        || node.NodeType?.Value == TimeNodeValues.AfterEffect);
                if (!uint.TryParse(target?.ShapeId?.Value,
                        NumberStyles.Integer, CultureInfo.InvariantCulture,
                        out uint shapeId)
                    || animations.FirstOrDefault(animation =>
                        animation.ShapeId == shapeId) is not
                        PowerPointClassicAnimation animation
                    || set.ChildElements.Count != 2
                    || behavior == null
                    || behavior.Descendants<ShapeTarget>().Count() != 1
                    || attributes.Length != 1
                    || !string.Equals(attributes[0].Text,
                        "style.visibility", StringComparison.Ordinal)
                    || value == null
                    || set.ToVariantValue!.ChildElements.Count != 1
                    || !string.Equals(value.Val?.Value, "visible",
                            StringComparison.OrdinalIgnoreCase)
                    && (!string.Equals(value.Val?.Value, "hidden",
                            StringComparison.OrdinalIgnoreCase)
                        || animation.AfterEffect is not
                            (PowerPointClassicAnimationAfterEffect
                                .HideImmediately or
                             PowerPointClassicAnimationAfterEffect
                                .HideOnNextClick))
                    || owner == null
                    || owner.Descendants<AnimateEffect>().Count(effect =>
                        uint.TryParse(effect.Descendants<ShapeTarget>()
                                .FirstOrDefault()?.ShapeId?.Value,
                            NumberStyles.Integer, CultureInfo.InvariantCulture,
                            out uint effectShapeId)
                        && effectShapeId == shapeId) != 1) {
                    return false;
                }
            }
            foreach (AnimateColor color in timing.Descendants<AnimateColor>()) {
                if (!animations.Any(animation =>
                        IsClassicDimAction(color, animation))) return false;
            }
            foreach (Audio audio in timing.Descendants<Audio>()) {
                if (!animations.Any(animation =>
                        IsClassicSoundAction(audio, animation))) return false;
            }
            foreach (Command command in timing.Descendants<Command>()) {
                if (!animations.Any(animation =>
                        IsClassicStopSoundAction(command, animation))) {
                    return false;
                }
            }
            foreach (PowerPointClassicAnimation animation in animations) {
                int hiddenCount = timing.Descendants<SetBehavior>().Count(set =>
                    IsClassicHiddenAction(set, animation));
                int dimCount = timing.Descendants<AnimateColor>().Count(color =>
                    IsClassicDimAction(color, animation));
                int soundCount = timing.Descendants<Audio>().Count(audio =>
                    IsClassicSoundAction(audio, animation));
                int stopSoundCount = timing.Descendants<Command>()
                    .Count(command => IsClassicStopSoundAction(command,
                        animation));
                bool expectsHidden = animation.AfterEffect is
                    PowerPointClassicAnimationAfterEffect.HideImmediately or
                    PowerPointClassicAnimationAfterEffect.HideOnNextClick;
                if (hiddenCount != (expectsHidden ? 1 : 0)
                    || dimCount != (animation.AfterEffect ==
                        PowerPointClassicAnimationAfterEffect.Dim ? 1 : 0)
                    || soundCount != (animation.PlaysSound
                        && !string.IsNullOrEmpty(
                            animation.SoundRelationshipId) ? 1 : 0)
                    || stopSoundCount != (animation.StopsSound ? 1 : 0)) {
                    return false;
                }
            }
            return true;
        }

        private static bool IsClassicHiddenAction(SetBehavior set,
            PowerPointClassicAnimation animation) =>
            IsClassicShapeBehavior(set.CommonBehavior, animation.ShapeId,
                "style.visibility")
            && string.Equals(set.ToVariantValue?
                    .GetFirstChild<StringVariantValue>()?.Val?.Value,
                "hidden", StringComparison.OrdinalIgnoreCase);

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
