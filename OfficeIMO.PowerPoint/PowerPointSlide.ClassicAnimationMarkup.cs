using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private static (OfficeXmlNode Timing, OfficeXmlNode Children, OfficeXmlNode Builds) CreateClassicTimingMarkup(
            OfficeXmlNode slide, IReadOnlyList<PowerPointClassicAnimation> animations, uint timingId) {
            const int p = ClassicPresentationNamespace;
            if ((ulong)timingId - 1 + CountClassicTimingIds(animations) > uint.MaxValue) throw TimingIdsExhausted();

            OfficeXmlNode timing = slide.CreateElement(p, "timing");
            OfficeXmlNode list = AppendTimingChild(timing, "tnLst");
            OfficeXmlNode root = AppendClassicTimingRoot(AppendTimingChild(list, "par"));
            OfficeXmlNode children = AppendTimingChild(root, "childTnLst");
            foreach (PowerPointClassicAnimation animation in animations) {
                AppendClassicEffect(children, animation, ref timingId);
            }
            return (timing, children, AppendClassicBuildList(timing, animations));
        }

        private static ulong CountClassicTimingIds(IReadOnlyList<PowerPointClassicAnimation> animations) {
            ulong count = 0;
            foreach (PowerPointClassicAnimation animation in animations) {
                count += 2UL;
                if (animation.StopsSound) count++;
                if (HasClassicSound(animation)) count++;
                if (animation.AfterEffect != PowerPointClassicAnimationAfterEffect.None) count++;
            }
            return count;
        }

        private static OfficeXmlNode AppendClassicTimingRoot(OfficeXmlNode parallel) {
            OfficeXmlNode root = AppendTimingChild(parallel, "cTn");
            root.SetUInt32Attribute("id", 1U);
            root.SetAttribute("dur", "indefinite");
            root.SetAttribute("restart", "never");
            root.SetAttribute("nodeType", "tmRoot");
            return root;
        }

        private static void AppendClassicEffect(OfficeXmlNode children, PowerPointClassicAnimation animation, ref uint timingId) {
            uint effectTimingId = timingId++;
            OfficeXmlNode owner = AppendTimingChild(AppendTimingChild(children, "par"), "cTn");
            OfficeXmlNode condition = CreateClassicCondition(owner, animation.Automatic ? animation.DelayMilliseconds : 0);
            if (!animation.Automatic) SetClassicClickCondition(condition);
            OfficeXmlNode actions = AppendTimingChild(owner, "childTnLst");
            OfficeXmlNode effect = AppendTimingChild(actions, "animEffect");
            effect.SetAttribute("transition", "in");
            effect.SetAttribute("filter", GetClassicAnimationFilter(animation.Effect, animation.Direction));
            OfficeXmlNode behavior = AppendTimingChild(effect, "cBhvr");
            OfficeXmlNode effectTime = AppendClassicActionTime(behavior, effectTimingId, "500");
            effectTime.SetAttribute("presetClass", "entr");
            effectTime.SetInt32Attribute("presetID", (byte)animation.Effect);
            effectTime.SetInt32Attribute("presetSubtype", animation.Direction);
            effectTime.SetBooleanAttribute("autoRev", animation.Reverse);
            effectTime.SetBooleanAttribute("afterEffect", animation.AfterEffect != PowerPointClassicAnimationAfterEffect.None);
            AppendClassicShapeTarget(behavior, animation.ShapeId);
            if (animation.StopsSound) AppendClassicStopSoundMarkup(actions, effectTimingId, timingId++);
            if (HasClassicSound(animation)) AppendClassicSoundMarkup(actions, animation, timingId++);
            if (animation.AfterEffect != PowerPointClassicAnimationAfterEffect.None) {
                AppendClassicAfterEffectMarkup(actions, animation, effectTimingId, timingId++);
            }
            owner.SetUInt32Attribute("id", timingId++);
            owner.SetAttribute("dur", "indefinite");
            owner.SetAttribute("fill", "hold");
            owner.SetAttribute("nodeType", animation.Automatic ? "afterEffect" : "clickEffect");
        }

        private static OfficeXmlNode AppendClassicActionTime(OfficeXmlNode container, uint timingId, string duration) {
            OfficeXmlNode actionTime = AppendTimingChild(container, "cTn");
            actionTime.SetUInt32Attribute("id", timingId);
            actionTime.SetAttribute("dur", duration);
            actionTime.SetAttribute("fill", "hold");
            return actionTime;
        }

        private static OfficeXmlNode CreateClassicCondition(OfficeXmlNode time, int delay) {
            OfficeXmlNode condition = AppendTimingChild(AppendTimingChild(time, "stCondLst"), "cond");
            condition.SetInt32Attribute("delay", delay);
            return condition;
        }

        private static void SetClassicClickCondition(OfficeXmlNode condition) {
            condition.SetAttribute("evt", "onClick");
            AppendTimingChild(AppendTimingChild(condition, "tgtEl"), "sldTgt");
        }

        private static void AppendClassicShapeTarget(OfficeXmlNode behavior, uint shapeId) {
            OfficeXmlNode target = AppendTimingChild(AppendTimingChild(behavior, "tgtEl"), "spTgt");
            target.SetUInt32Attribute("spid", shapeId);
        }

        private static void AppendClassicStopSoundMarkup(OfficeXmlNode actions, uint effectTimingId, uint timingId) {
            OfficeXmlNode command = AppendTimingChild(actions, "cmd");
            command.SetAttribute("type", "evt");
            command.SetAttribute("cmd", "onstopaudio");
            OfficeXmlNode behavior = AppendTimingChild(command, "cBhvr");
            OfficeXmlNode time = AppendClassicActionTime(behavior, timingId, "1");
            time.SetBooleanAttribute("display", false);
            time.SetAttribute("masterRel", "sameClick");
            OfficeXmlNode condition = CreateClassicCondition(time, 0);
            condition.SetAttribute("evt", "begin");
            AppendTimingChild(condition, "tn").SetUInt32Attribute("val", effectTimingId);
            AppendTimingChild(AppendTimingChild(behavior, "tgtEl"), "sldTgt");
        }

        private static void AppendClassicSoundMarkup(OfficeXmlNode actions, PowerPointClassicAnimation animation, uint timingId) {
            OfficeXmlNode media = AppendTimingChild(AppendTimingChild(actions, "audio"), "cMediaNode");
            media.SetInt32Attribute("vol", 80000);
            CreateClassicCondition(AppendClassicActionTime(media, timingId, "media"), 0);
            OfficeXmlNode sound = AppendTimingChild(AppendTimingChild(media, "tgtEl"), "sndTgt");
            sound.SetAttribute(OfficeXmlNamespaces.Relationships, "embed", animation.SoundRelationshipId);
            sound.SetAttribute("name", animation.SoundName ?? "Animation Sound");
        }

        private static bool HasClassicSound(PowerPointClassicAnimation animation) =>
            animation.PlaysSound && !string.IsNullOrEmpty(animation.SoundRelationshipId);

        private static void AppendClassicAfterEffectMarkup(OfficeXmlNode actions, PowerPointClassicAnimation animation,
            uint effectTimingId, uint timingId) {
            bool dim = animation.AfterEffect == PowerPointClassicAnimationAfterEffect.Dim;
            OfficeXmlNode action = AppendTimingChild(actions, dim ? "animClr" : "set");
            OfficeXmlNode behavior = AppendTimingChild(action, "cBhvr");
            OfficeXmlNode condition = CreateClassicCondition(AppendClassicActionTime(behavior, timingId, "1"), 0);
            if (animation.AfterEffect == PowerPointClassicAnimationAfterEffect.HideOnNextClick) {
                SetClassicClickCondition(condition);
            } else {
                condition.SetAttribute("evt", "end");
                AppendTimingChild(condition, "tn").SetUInt32Attribute("val", effectTimingId);
            }
            AppendClassicShapeTarget(behavior, animation.ShapeId);
            OfficeXmlNode attribute = AppendTimingChild(AppendTimingChild(behavior, "attrNameLst"), "attrName");
            attribute.SetInnerText(dim ? "style.color" : "style.visibility");
            OfficeXmlNode to = AppendTimingChild(action, "to");
            if (dim) AppendClassicDimColor(to, animation.RawDimColor);
            else AppendTimingChild(to, "strVal").SetAttribute("val", "hidden");
        }

        private static void AppendClassicDimColor(OfficeXmlNode to, uint rawColor) {
            const int a = OfficeXmlNamespaces.DrawingML;
            string? scheme = (rawColor >> 24) switch {
                0 => "lt1", 1 => "dk1", 2 => "accent4", 3 => "dk2",
                4 => "lt2", 5 => "accent1", 6 => "accent2", 7 => "accent3",
                _ => null
            };
            if (scheme != null) {
                to.Append(to.CreateElement(a, "schemeClr")).SetAttribute("val", scheme);
                return;
            }
            string rgb = string.Concat(
                unchecked((byte)rawColor).ToString("X2", CultureInfo.InvariantCulture),
                unchecked((byte)(rawColor >> 8)).ToString("X2", CultureInfo.InvariantCulture),
                unchecked((byte)(rawColor >> 16)).ToString("X2", CultureInfo.InvariantCulture));
            to.Append(to.CreateElement(a, "srgbClr")).SetAttribute("val", rgb);
        }

        private static OfficeXmlNode AppendClassicBuildList(OfficeXmlNode timing, IEnumerable<PowerPointClassicAnimation> animations) {
            OfficeXmlNode list = timing.CreateElement(ClassicPresentationNamespace, "bldLst");
            foreach (PowerPointClassicAnimation animation in animations) {
                bool paragraph = animation.BuildType is >= PowerPointClassicAnimationBuildType.ByLevel1Paragraph
                    and <= PowerPointClassicAnimationBuildType.ByLevel5Paragraph;
                if (!paragraph && animation.BuildType != PowerPointClassicAnimationBuildType.AsOneObject) continue;
                OfficeXmlNode build = AppendTimingChild(list, "bldP");
                build.SetUInt32Attribute("spid", animation.ShapeId);
                build.SetUInt32Attribute("grpId", 0);
                build.SetAttribute("build", paragraph ? "p" : "whole");
                build.SetBooleanAttribute("animBg", animation.AnimateBackground);
                if (!paragraph) continue;
                build.SetUInt32Attribute("bldLvl", (uint)((byte)animation.BuildType - 1));
                build.SetBooleanAttribute("rev", animation.Reverse);
                if (animation.Automatic) build.SetInt32Attribute("advAuto", animation.DelayMilliseconds);
            }
            return list.FirstChild.IsNone ? default : timing.Append(list);
        }

        private static void WriteClassicAnimationMetadata(OfficeXmlNode slide, IReadOnlyList<PowerPointClassicAnimation> animations) {
            const int p = ClassicPresentationNamespace;
            OfficeXmlNode list = slide.Child(p, "extLst");
            foreach (OfficeXmlNode item in list.Children(p, "ext")) {
                if (item.Attribute("uri") != ClassicAnimationExtensionUri) continue;
                item.Remove();
                break;
            }
            if (animations.Count == 0) {
                if (!ClassicChildren(list).Any()) list.Remove();
                return;
            }
            if (list.IsNone) list = AppendTimingChild(slide, "extLst");
            OfficeXmlNode extension = AppendTimingChild(list, "ext");
            extension.SetAttribute("uri", ClassicAnimationExtensionUri);
            OfficeXmlNode metadata = extension.Append(extension.CreateElement(ClassicAnimationNamespace, "classicAnimations", "oimo"));
            foreach (PowerPointClassicAnimation animation in animations) {
                OfficeXmlNode element = metadata.Append(metadata.CreateElement(ClassicAnimationNamespace, "animation", "oimo"));
                element.SetUInt32Attribute("shapeId", animation.ShapeId);
                element.SetInt32Attribute("effect", (byte)animation.Effect);
                element.SetInt32Attribute("direction", animation.Direction);
                element.SetInt32Attribute("build", (byte)animation.BuildType);
                element.SetAttribute("automatic", animation.Automatic.ToString());
                element.SetInt32Attribute("delay", animation.DelayMilliseconds);
                element.SetInt32Attribute("order", animation.Order);
                element.SetAttribute("reverse", animation.Reverse.ToString());
                element.SetAttribute("animateBackground", animation.AnimateBackground.ToString());
                element.SetInt32Attribute("after", (byte)animation.AfterEffect);
                element.SetInt32Attribute("textBuild", (byte)animation.TextBuild);
                element.SetUInt32Attribute("dimColor", animation.RawDimColor);
                element.SetAttribute("playsSound", animation.PlaysSound.ToString());
                element.SetAttribute("stopsSound", animation.StopsSound.ToString());
                if (!HasClassicSound(animation)) continue;
                OfficeXmlNode sound = AppendTimingChild(element, "snd");
                sound.SetAttribute(OfficeXmlNamespaces.Relationships, "embed", animation.SoundRelationshipId);
                sound.SetAttribute("name", animation.SoundName ?? "Animation Sound");
            }
        }
    }
}
