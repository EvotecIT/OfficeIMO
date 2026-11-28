using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private const int TransitionPresentationNamespace = OfficeXmlNamespaces.PresentationML;
        private const int TransitionCompatibilityNamespace = OfficeXmlNamespaces.MarkupCompatibility;
        private const string PowerPoint2012Namespace = "http://schemas.microsoft.com/office/powerpoint/2012/main";

        /// <summary>Transition applied when moving to this slide.</summary>
        public PowerPointSlideTransition Transition {
            get => GetOwnedTransitionValue(GetTransitionNode());
            set {
                PowerPointPartRef part = PartRef.ForMovingXmlNodes();
                OfficeXmlNode root = part.Root;
                OfficeXmlNode previous = GetTransitionNode(root);
                PowerPointSlideTransitionSpeed? speed = ReadTransitionSpeed(previous);
                double? duration = ReadTransitionSeconds(previous.Attribute(P14Namespace, "dur"));
                bool? advanceOnClick = previous.BooleanAttribute("advClick");
                double? advanceAfter = ReadTransitionSeconds(previous.Attribute("advTm"));
                OfficeXmlNode sound = previous.Child(TransitionPresentationNamespace, "sndAc").Clone();
                string[] soundIds = GetTransitionSoundRelationshipIds();

                root.Child(TransitionPresentationNamespace, "transition").Remove();
                GetTransitionAlternate(root).Remove();

                if (value == PowerPointSlideTransition.Morph) {
                    WriteMorphTransition(root);
                } else if (value != PowerPointSlideTransition.None) {
                    OfficeXmlNode transition = root.GetOrAddChild(TransitionPresentationNamespace, "transition", PSlideView.Order);
                    OfficeXmlNode effect = CreateClassicTransition(transition, value);
                    if (effect.IsNone) effect = CreatePowerPoint2010Transition(transition, value);
                    if (effect.Exists) transition.Append(effect);
                }

                foreach (OfficeXmlNode transition in GetTransitionNodes(root)) {
                    if (sound.Exists) transition.Append(sound.Clone());
                    WriteTransitionSpeed(transition, speed);
                    WriteTransitionDuration(transition, duration);
                    transition.SetBooleanAttribute("advClick", advanceOnClick);
                    transition.SetAttribute("advTm", ToMillisecondsString(advanceAfter));
                }

                part.MarkEdited();
                RemoveUnusedSlideSounds(soundIds);
            }
        }

        private static PowerPointSlideTransition GetOwnedTransitionValue(OfficeXmlNode transition) {
            if (transition.IsNone) return PowerPointSlideTransition.None;
            if (GetClassicTransition(transition) is { } classic) return classic;
            if (transition.Child(P14Namespace, "flash").Exists) return PowerPointSlideTransition.Flash;

            OfficeXmlNode warp = transition.Child(P14Namespace, "warp");
            if (warp.Exists) return warp.Attribute("dir") == "out" ? PowerPointSlideTransition.WarpOut : PowerPointSlideTransition.WarpIn;
            if (transition.Child(P14Namespace, "prism").Exists) return PowerPointSlideTransition.Prism;

            OfficeXmlNode ferris = transition.Child(P14Namespace, "ferris");
            if (ferris.Exists) return ferris.Attribute("dir") == "r" ? PowerPointSlideTransition.FerrisRight : PowerPointSlideTransition.FerrisLeft;
            if (transition.Child(P159Namespace, "morph").Exists) return PowerPointSlideTransition.Morph;

            return string.Equals(transition.Child(PowerPoint2012Namespace, "prstTrans").Attribute("prst"), "morph", StringComparison.OrdinalIgnoreCase)
                ? PowerPointSlideTransition.Morph : PowerPointSlideTransition.None;
        }

        private static OfficeXmlNode CreatePowerPoint2010Transition(OfficeXmlNode transition, PowerPointSlideTransition value) {
            (string? name, string? attribute, string? attributeValue) = value switch {
                PowerPointSlideTransition.Flash => ("flash", null, null),
                PowerPointSlideTransition.WarpIn => ("warp", "dir", "in"),
                PowerPointSlideTransition.WarpOut => ("warp", "dir", "out"),
                PowerPointSlideTransition.Prism => ("prism", "isContent", "1"),
                PowerPointSlideTransition.FerrisLeft => ("ferris", "dir", "l"),
                PowerPointSlideTransition.FerrisRight => ("ferris", "dir", "r"),
                _ => default
            };
            if (name == null) return default;
            transition.DeclareNamespace("p14", P14Namespace);
            OfficeXmlNode effect = transition.CreateElement(P14Namespace, name, "p14");
            if (attribute != null) effect.SetAttribute(attribute, attributeValue);
            return effect;
        }

        /// <summary>Gets or sets the optional transition playback speed.</summary>
        public PowerPointSlideTransitionSpeed? TransitionSpeed {
            get => ReadTransitionSpeed(GetTransitionNode());
            set {
                PowerPointPartRef part = PartRef;
                foreach (OfficeXmlNode transition in GetTransitionNodes(part.Root)) {
                    WriteTransitionSpeed(transition, value);
                    part.MarkEdited();
                }
            }
        }

        /// <summary>Gets or sets the optional transition duration in seconds.</summary>
        public double? TransitionDurationSeconds {
            get => ReadTransitionSeconds(GetTransitionNode().Attribute(P14Namespace, "dur"));
            set {
                PowerPointPartRef part = PartRef;
                foreach (OfficeXmlNode transition in GetTransitionNodes(part.Root)) {
                    WriteTransitionDuration(transition, value);
                    part.MarkEdited();
                }
            }
        }

        /// <summary>Gets or sets whether clicking advances the slide.</summary>
        public bool? TransitionAdvanceOnClick {
            get => GetTransitionNode().BooleanAttribute("advClick");
            set {
                PowerPointPartRef part = PartRef;
                foreach (OfficeXmlNode transition in GetTransitionNodes(part.Root)) {
                    transition.SetBooleanAttribute("advClick", value);
                    part.MarkEdited();
                }
            }
        }

        /// <summary>Gets or sets the optional automatic-advance time in seconds.</summary>
        public double? TransitionAdvanceAfterSeconds {
            get => ReadTransitionSeconds(GetTransitionNode().Attribute("advTm"));
            set {
                PowerPointPartRef part = PartRef;
                foreach (OfficeXmlNode transition in GetTransitionNodes(part.Root)) {
                    transition.SetAttribute("advTm", ToMillisecondsString(value));
                    part.MarkEdited();
                }
            }
        }

        private static PowerPointSlideTransitionSpeed? ReadTransitionSpeed(OfficeXmlNode node) => node.Attribute("spd") switch {
            "slow" => PowerPointSlideTransitionSpeed.Slow,
            "fast" => PowerPointSlideTransitionSpeed.Fast,
            "med" => PowerPointSlideTransitionSpeed.Medium,
            _ => null
        };

        private static void WriteTransitionSpeed(OfficeXmlNode node, PowerPointSlideTransitionSpeed? speed) =>
            node.SetAttribute("spd", speed switch {
                PowerPointSlideTransitionSpeed.Slow => "slow",
                PowerPointSlideTransitionSpeed.Fast => "fast",
                PowerPointSlideTransitionSpeed.Medium => "med",
                _ => null
            });

        private static double? ReadTransitionSeconds(string? value) => uint.TryParse(value, out uint milliseconds)
            ? milliseconds / 1000.0 : null;

        private static string? ToMillisecondsString(double? seconds) {
            if (!seconds.HasValue) return null;
            uint milliseconds = seconds.Value < 0 ? 0 : (uint)Math.Round(seconds.Value * 1000.0, MidpointRounding.AwayFromZero);
            return milliseconds.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static void WriteTransitionDuration(OfficeXmlNode transition, double? seconds) {
            if (seconds.HasValue) {
                EnsureTransitionNamespace(transition, "p14", P14Namespace);
                OfficeXmlNode slide = transition.AncestorOrSelf(TransitionPresentationNamespace, "sld");
                if (slide.Exists) MarkNamespaceIgnorable(slide, "p14", P14Namespace);
            }
            transition.SetAttribute(P14Namespace, "dur", ToMillisecondsString(seconds), "p14");
        }

        private OfficeXmlNode GetTransitionNode() => GetTransitionNode(PartRef.Root);

        private static OfficeXmlNode GetTransitionNode(OfficeXmlNode root) {
            OfficeXmlNode direct = root.Child(TransitionPresentationNamespace, "transition");
            if (direct.Exists) return direct;
            foreach (OfficeXmlNode branch in GetAlternateTransitionNodes(GetTransitionAlternate(root))) {
                if (branch.Exists) return branch;
            }
            return default;
        }

        private static IEnumerable<OfficeXmlNode> GetTransitionNodes(OfficeXmlNode root) {
            OfficeXmlNode direct = root.Child(TransitionPresentationNamespace, "transition");
            if (direct.Exists) {
                yield return direct;
                yield break;
            }
            foreach (OfficeXmlNode branch in GetAlternateTransitionNodes(GetTransitionAlternate(root))) {
                if (branch.Exists) yield return branch;
            }
        }

        private static OfficeXmlNode GetTransitionAlternate(OfficeXmlNode root) {
            foreach (OfficeXmlNode alternate in root.Children(TransitionCompatibilityNamespace, "AlternateContent")) {
                if (GetAlternateTransitionNodes(alternate).Any(branch => branch.Exists)) return alternate;
            }
            return default;
        }

        private static IEnumerable<OfficeXmlNode> GetAlternateTransitionNodes(OfficeXmlNode alternate) {
            foreach (OfficeXmlNode choice in alternate.Children(TransitionCompatibilityNamespace, "Choice")) {
                yield return choice.Child(TransitionPresentationNamespace, "transition");
            }
            OfficeXmlNode fallback = alternate.Child(TransitionCompatibilityNamespace, "Fallback");
            if (fallback.Exists) yield return fallback.Child(TransitionPresentationNamespace, "transition");
        }

        private static void WriteMorphTransition(OfficeXmlNode slide) {
            MarkNamespaceIgnorable(slide, "p159", P159Namespace);
            OfficeXmlNode alternate = slide.CreateElement(TransitionCompatibilityNamespace, "AlternateContent");
            OfficeXmlNode choice = alternate.Append(alternate.CreateElement(TransitionCompatibilityNamespace, "Choice"));
            choice.SetAttribute("Requires", "p159");
            // A detached compatibility branch has no slide type information in the SDK.
            OfficeXmlNode transition = choice.Append(slide.CreateElement(TransitionPresentationNamespace, "transition"));
            transition.DeclareNamespace("p159", P159Namespace);
            OfficeXmlNode morph = transition.Append(transition.CreateElement(P159Namespace, "morph", "p159"));
            morph.DeclareNamespace("p159", P159Namespace);
            morph.SetAttribute("option", "byObject");

            OfficeXmlNode fallback = alternate.Append(alternate.CreateElement(TransitionCompatibilityNamespace, "Fallback"));
            OfficeXmlNode alternative = fallback.Append(slide.CreateElement(TransitionPresentationNamespace, "transition"));
            alternative.Append(alternative.CreateElement(TransitionPresentationNamespace, "fade"));

            OfficeXmlNode before = slide.Child(TransitionPresentationNamespace, "timing");
            if (before.IsNone) before = slide.Child(TransitionPresentationNamespace, "extLst");
            if (before.IsNone) {
                OfficeXmlNode after = slide.Child(TransitionPresentationNamespace, "clrMapOvr");
                if (after.IsNone) after = slide.Child(TransitionPresentationNamespace, "cSld");
                before = after.NextNode();
            }
            slide.InsertBefore(alternate, before);
        }

        private static void EnsureTransitionNamespace(OfficeXmlNode node, string prefix, string uri) {
            if (!string.Equals(node.LookupNamespace(prefix), uri, StringComparison.Ordinal)) node.DeclareNamespace(prefix, uri);
        }

        private static void MarkNamespaceIgnorable(OfficeXmlNode slide, string prefix, string uri) {
            EnsureTransitionNamespace(slide, "mc", MarkupCompatibilityNamespace);
            EnsureTransitionNamespace(slide, prefix, uri);
            List<string> ignorable = (slide.Attribute(TransitionCompatibilityNamespace, "Ignorable") ?? string.Empty)
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            if (!ignorable.Contains(prefix, StringComparer.Ordinal)) ignorable.Add(prefix);
            slide.SetAttribute(TransitionCompatibilityNamespace, "Ignorable", string.Join(" ", ignorable));
        }
    }
}
