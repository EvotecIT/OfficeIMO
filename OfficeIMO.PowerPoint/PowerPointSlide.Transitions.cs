using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Transition applied when moving to this slide.
        /// </summary>
        public SlideTransition Transition {
            get {
                Transition? t = GetTransitionElement();
                if (t == null) {
                    return SlideTransition.None;
                }

                SlideTransition? classicTransition = GetClassicTransition(t);
                if (classicTransition.HasValue) {
                    return classicTransition.Value;
                }

                if (t.GetFirstChild<P14.FlashTransition>() != null) {
                    return SlideTransition.Flash;
                }

                P14.WarpTransition? warp = t.GetFirstChild<P14.WarpTransition>();
                if (warp != null) {
                    return warp.Direction?.Value == TransitionInOutDirectionValues.Out
                        ? SlideTransition.WarpOut
                        : SlideTransition.WarpIn;
                }

                if (t.GetFirstChild<P14.PrismTransition>() != null) {
                    return SlideTransition.Prism;
                }

                P14.FerrisTransition? ferris = t.GetFirstChild<P14.FerrisTransition>();
                if (ferris != null) {
                    return ferris.Direction?.Value == P14.TransitionLeftRightDirectionTypeValues.Right
                        ? SlideTransition.FerrisRight
                        : SlideTransition.FerrisLeft;
                }

                if (HasMorphTransition(t)) {
                    return SlideTransition.Morph;
                }

                return SlideTransition.None;
            }
            set {
                SlideTransitionSpeed? speed = TransitionSpeed;
                double? durationSeconds = TransitionDurationSeconds;
                bool? advanceOnClick = TransitionAdvanceOnClick;
                double? advanceAfterSeconds = TransitionAdvanceAfterSeconds;
                SoundAction? soundAction = GetTransitionElement()?
                    .GetFirstChild<SoundAction>()?.CloneNode(true) as SoundAction;
                string[] soundRelationshipIds =
                    GetTransitionSoundRelationshipIds();

                RemoveTransitionMarkup();
                if (value == SlideTransition.None) {
                    RemoveUnusedTransitionSounds(soundRelationshipIds);
                    return;
                }

                if (value == SlideTransition.Morph) {
                    SetMorphTransition();
                    foreach (Transition morphTransition in
                             GetTransitionElements()) {
                        if (soundAction != null) {
                            morphTransition.Append(soundAction.CloneNode(true));
                        }
                        ApplyTransitionSettings(morphTransition, speed,
                            durationSeconds, advanceOnClick,
                            advanceAfterSeconds);
                    }
                    RemoveUnusedTransitionSounds(soundRelationshipIds);
                    return;
                }

                Transition transition = new();
                OpenXmlElement? classicTransition = CreateClassicTransition(value);
                if (classicTransition != null) {
                    transition.Append(classicTransition);
                } else {
                    switch (value) {
                        case SlideTransition.Flash:
                            transition.AddNamespaceDeclaration("p14", P14Namespace);
                            transition.Append(new P14.FlashTransition());
                            break;
                        case SlideTransition.WarpIn:
                            transition.AddNamespaceDeclaration("p14", P14Namespace);
                            transition.Append(new P14.WarpTransition { Direction = TransitionInOutDirectionValues.In });
                            break;
                        case SlideTransition.WarpOut:
                            transition.AddNamespaceDeclaration("p14", P14Namespace);
                            transition.Append(new P14.WarpTransition { Direction = TransitionInOutDirectionValues.Out });
                            break;
                        case SlideTransition.Prism:
                            transition.AddNamespaceDeclaration("p14", P14Namespace);
                            transition.Append(new P14.PrismTransition { IsContent = true });
                            break;
                        case SlideTransition.FerrisLeft:
                            transition.AddNamespaceDeclaration("p14", P14Namespace);
                            transition.Append(new P14.FerrisTransition { Direction = P14.TransitionLeftRightDirectionTypeValues.Left });
                            break;
                        case SlideTransition.FerrisRight:
                            transition.AddNamespaceDeclaration("p14", P14Namespace);
                            transition.Append(new P14.FerrisTransition { Direction = P14.TransitionLeftRightDirectionTypeValues.Right });
                            break;
                    }
                }

                SlideRoot.Transition = transition;
                if (soundAction != null) transition.Append(soundAction);
                ApplyTransitionSettings(GetTransitionElement(), speed, durationSeconds, advanceOnClick, advanceAfterSeconds);
                RemoveUnusedTransitionSounds(soundRelationshipIds);
            }
        }

        private void RemoveUnusedTransitionSounds(
            IEnumerable<string> relationshipIds) {
            foreach (string relationshipId in relationshipIds) {
                PowerPointEmbeddedSound.RemoveIfUnused(_slidePart,
                    relationshipId);
            }
        }

        /// <summary>
        ///     Gets or sets the optional transition playback speed.
        /// </summary>
        public SlideTransitionSpeed? TransitionSpeed {
            get {
                Transition? transition = GetTransitionElement();
                if (transition?.Speed?.Value == null) {
                    return null;
                }

                var speedValue = transition.Speed.Value;
                if (speedValue == TransitionSpeedValues.Slow) {
                    return SlideTransitionSpeed.Slow;
                }

                if (speedValue == TransitionSpeedValues.Fast) {
                    return SlideTransitionSpeed.Fast;
                }

                return SlideTransitionSpeed.Medium;
            }
            set {
                foreach (Transition transition in GetTransitionElements()) {
                    transition.Speed = value switch {
                        SlideTransitionSpeed.Slow =>
                            TransitionSpeedValues.Slow,
                        SlideTransitionSpeed.Fast =>
                            TransitionSpeedValues.Fast,
                        SlideTransitionSpeed.Medium =>
                            TransitionSpeedValues.Medium,
                        _ => null
                    };
                }
            }
        }

        /// <summary>
        ///     Gets or sets the optional transition duration in seconds.
        /// </summary>
        public double? TransitionDurationSeconds {
            get {
                uint? milliseconds = ParseTransitionMilliseconds(GetTransitionElement()?.Duration?.Value);
                if (milliseconds == null) {
                    return null;
                }

                return milliseconds.Value / 1000.0;
            }
            set {
                foreach (Transition transition in GetTransitionElements()) {
                    if (value.HasValue) {
                        EnsureTransitionCompatibilityNamespace(transition,
                            "p14", P14Namespace);
                        transition.Duration = ToMillisecondsString(value);
                    } else {
                        RemoveTransitionAttribute(transition, "dur",
                            P14Namespace);
                    }
                }
            }
        }

        /// <summary>
        ///     Gets or sets whether clicking advances the slide.
        /// </summary>
        public bool? TransitionAdvanceOnClick {
            get {
                return GetTransitionElement()?.AdvanceOnClick?.Value;
            }
            set {
                foreach (Transition transition in GetTransitionElements()) {
                    if (value.HasValue) {
                        transition.AdvanceOnClick = value;
                    } else {
                        RemoveTransitionAttribute(transition, "advClick",
                            string.Empty);
                    }
                }
            }
        }

        /// <summary>
        ///     Gets or sets the optional automatic-advance time in seconds.
        /// </summary>
        public double? TransitionAdvanceAfterSeconds {
            get {
                uint? milliseconds = ParseTransitionMilliseconds(GetTransitionElement()?.AdvanceAfterTime?.Value);
                if (milliseconds == null) {
                    return null;
                }

                return milliseconds.Value / 1000.0;
            }
            set {
                foreach (Transition transition in GetTransitionElements()) {
                    if (value.HasValue) {
                        transition.AdvanceAfterTime =
                            ToMillisecondsString(value);
                    } else {
                        RemoveTransitionAttribute(transition, "advTm",
                            string.Empty);
                    }
                }
            }
        }

        private static bool HasMorphTransition(Transition transition) {
            return transition.ChildElements.Any(element =>
                (element.LocalName == "morph" && element.NamespaceUri == P159Namespace) ||
                (element.LocalName == "prstTrans" &&
                 element.NamespaceUri == "http://schemas.microsoft.com/office/powerpoint/2012/main" &&
                 string.Equals(element.GetAttribute("prst", string.Empty).Value, "morph", StringComparison.OrdinalIgnoreCase)));
        }

        private static OpenXmlUnknownElement CreateMorphTransition() {
            OpenXmlUnknownElement morph = new OpenXmlUnknownElement("p159", "morph", P159Namespace);
            morph.AddNamespaceDeclaration("p159", P159Namespace);
            morph.SetAttribute(new OpenXmlAttribute("option", string.Empty, "byObject"));
            return morph;
        }

        private static uint? ToMilliseconds(double? seconds) {
            if (!seconds.HasValue) {
                return null;
            }

            if (seconds.Value < 0) {
                return 0;
            }

            return (uint)Math.Round(seconds.Value * 1000.0, MidpointRounding.AwayFromZero);
        }

        private static string? ToMillisecondsString(double? seconds) {
            uint? milliseconds = ToMilliseconds(seconds);
            return milliseconds?.ToString();
        }

        private static uint? ParseTransitionMilliseconds(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            return uint.TryParse(value, out uint parsed)
                ? parsed
                : null;
        }

        private static void ApplyTransitionSettings(Transition? transition, SlideTransitionSpeed? speed, double? durationSeconds, bool? advanceOnClick, double? advanceAfterSeconds) {
            if (transition == null) {
                return;
            }

            transition.Speed = speed switch {
                SlideTransitionSpeed.Slow => TransitionSpeedValues.Slow,
                SlideTransitionSpeed.Fast => TransitionSpeedValues.Fast,
                SlideTransitionSpeed.Medium => TransitionSpeedValues.Medium,
                _ => null
            };

            if (durationSeconds.HasValue) {
                EnsureTransitionCompatibilityNamespace(transition, "p14", P14Namespace);
                transition.Duration = ToMillisecondsString(durationSeconds);
            } else {
                RemoveTransitionAttribute(transition, "dur", P14Namespace);
            }

            if (advanceOnClick.HasValue) {
                transition.AdvanceOnClick = advanceOnClick;
            } else {
                RemoveTransitionAttribute(transition, "advClick", string.Empty);
            }

            if (advanceAfterSeconds.HasValue) {
                transition.AdvanceAfterTime = ToMillisecondsString(advanceAfterSeconds);
            } else {
                RemoveTransitionAttribute(transition, "advTm", string.Empty);
            }
        }

        private static void EnsureTransitionCompatibilityNamespace(Transition transition, string prefix, string uri) {
            if (!string.Equals(transition.LookupNamespace(prefix), uri, StringComparison.Ordinal)) {
                transition.AddNamespaceDeclaration(prefix, uri);
            }

            Slide? slide = transition.Ancestors<Slide>().FirstOrDefault();
            if (slide == null) {
                return;
            }

            if (!string.Equals(slide.LookupNamespace("mc"), MarkupCompatibilityNamespace, StringComparison.Ordinal)) {
                slide.AddNamespaceDeclaration("mc", MarkupCompatibilityNamespace);
            }

            if (!string.Equals(slide.LookupNamespace(prefix), uri, StringComparison.Ordinal)) {
                slide.AddNamespaceDeclaration(prefix, uri);
            }

            slide.MCAttributes = MergeIgnorableNamespace(slide.MCAttributes, prefix);
        }

        private static void RemoveTransitionAttribute(Transition transition, string localName, string namespaceUri) {
            bool hasAttribute = transition.GetAttributes()
                .Any(attribute =>
                    string.Equals(attribute.LocalName, localName, StringComparison.Ordinal) &&
                    string.Equals(attribute.NamespaceUri, namespaceUri, StringComparison.Ordinal));

            if (hasAttribute) {
                transition.RemoveAttribute(localName, namespaceUri);
            }
        }

        internal Transition? GetTransitionElement() {
            return GetTransitionElements().FirstOrDefault();
        }

        internal IReadOnlyList<Transition> GetTransitionElements() {
            if (SlideRoot.Transition is Transition transition) {
                return new[] { transition };
            }
            AlternateContent? alternateContent = GetTransitionAlternateContent();
            if (alternateContent == null) {
                return Array.Empty<Transition>();
            }
            return alternateContent.Elements<AlternateContentChoice>()
                .Select(choice => choice.GetFirstChild<Transition>())
                .Concat(new[] { alternateContent
                    .GetFirstChild<AlternateContentFallback>()?
                    .GetFirstChild<Transition>() })
                .Where(candidate => candidate != null)
                .Cast<Transition>()
                .ToArray();
        }

        private AlternateContent? GetTransitionAlternateContent() {
            return SlideRoot.Elements<AlternateContent>()
                .FirstOrDefault(content =>
                    content.Elements<AlternateContentChoice>().Any(choice => choice.GetFirstChild<Transition>() != null) ||
                    content.GetFirstChild<AlternateContentFallback>()?.GetFirstChild<Transition>() != null);
        }

        private void RemoveTransitionMarkup() {
            SlideRoot.Transition = null;
            GetTransitionAlternateContent()?.Remove();
        }

        private void SetMorphTransition() {
            Slide slide = SlideRoot;
            if (!string.Equals(slide.LookupNamespace("mc"),
                    MarkupCompatibilityNamespace, StringComparison.Ordinal)) {
                slide.AddNamespaceDeclaration("mc",
                    MarkupCompatibilityNamespace);
            }
            if (!string.Equals(slide.LookupNamespace("p159"),
                    P159Namespace, StringComparison.Ordinal)) {
                slide.AddNamespaceDeclaration("p159", P159Namespace);
            }
            slide.MCAttributes = MergeIgnorableNamespace(slide.MCAttributes, "p159");

            Transition morphTransition = new();
            morphTransition.AddNamespaceDeclaration("p159", P159Namespace);
            morphTransition.Append(CreateMorphTransition());

            AlternateContentChoice choice = new() { Requires = "p159" };
            choice.Append(morphTransition);

            AlternateContentFallback fallback = new();
            fallback.Append(new Transition(new FadeTransition()));

            AlternateContent alternateContent = new();
            alternateContent.Append(choice);
            alternateContent.Append(fallback);

            InsertTransitionAlternateContent(alternateContent);
        }

        private void InsertTransitionAlternateContent(AlternateContent alternateContent) {
            Slide slide = SlideRoot;
            OpenXmlElement? insertBefore = slide.GetFirstChild<Timing>();
            insertBefore ??= slide.GetFirstChild<ExtensionListWithModification>();
            if (insertBefore != null) {
                slide.InsertBefore(alternateContent, insertBefore);
                return;
            }

            if (slide.ColorMapOverride != null) {
                slide.InsertAfter(alternateContent, slide.ColorMapOverride);
                return;
            }

            if (slide.CommonSlideData != null) {
                slide.InsertAfter(alternateContent, slide.CommonSlideData);
                return;
            }

            slide.Append(alternateContent);
        }

        private static MarkupCompatibilityAttributes MergeIgnorableNamespace(
            MarkupCompatibilityAttributes? existingAttributes,
            string namespacePrefix) {
            List<string> prefixes = (existingAttributes?.Ignorable?.Value ?? string.Empty)
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .ToList();

            if (!prefixes.Contains(namespacePrefix, StringComparer.Ordinal)) {
                prefixes.Add(namespacePrefix);
            }

            return new MarkupCompatibilityAttributes {
                Ignorable = string.Join(" ", prefixes)
            };
        }
    }
}
