using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private const string TransitionAudioRelationship = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "audio";

        // The legacy binary writer reads each saved branch and rejects incompatible branch settings.
        internal IReadOnlyList<Transition?> GetTransitionBranches() {
            OfficeXmlNode root = OfficeXmlNode.FromSdk(SlideRoot);
            OfficeXmlNode primary = root.Child(TransitionPresentationNamespace, "transition");
            if (primary.SdkElement is Transition transition) return new Transition?[] { transition };
            return GetAlternateTransitionNodes(GetTransitionAlternate(root))
                .Select(branch => branch.SdkElement as Transition).ToArray();
        }

        internal static PowerPointSlideTransition GetTransitionValue(Transition transition) =>
            GetOwnedTransitionValue(OfficeXmlNode.FromSdk(transition));

        /// <summary>Gets whether this slide starts an embedded transition sound.</summary>
        public bool HasTransitionSound => TransitionSound(GetTransitionNode()).Exists;

        /// <summary>Gets the embedded transition sound name, when present.</summary>
        public string? TransitionSoundName => TransitionSound(GetTransitionNode()).Attribute("name");

        /// <summary>Gets whether the embedded transition sound loops.</summary>
        public bool TransitionSoundLoops => TransitionStartSound(GetTransitionNode()).BooleanAttribute("loop") == true;

        /// <summary>Gets whether this transition stops the currently playing sound.</summary>
        public bool TransitionStopsSound => GetTransitionNode().Child(TransitionPresentationNamespace, "sndAc")
            .Child(TransitionPresentationNamespace, "endSnd").Exists;

        /// <summary>Sets an embedded WAV or AIFF transition sound from a file.</summary>
        public void SetTransitionSound(string audioPath, bool loop = false) {
            ArgumentNullException.ThrowIfNull(audioPath);
            if (!File.Exists(audioPath)) throw new FileNotFoundException("Audio file not found.", audioPath);
            using FileStream input = new(audioPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            SetTransitionSound(input, Path.GetFileName(audioPath), GetAudioContentType(audioPath), Path.GetExtension(audioPath), loop);
        }

        /// <summary>Sets an embedded WAV or AIFF transition sound.</summary>
        public void SetTransitionSound(Stream audio, string name, string contentType = "audio/wav",
            string extension = ".wav", bool loop = false) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("A transition sound name is required.", nameof(name));
            }
            PowerPointPartRef part = PartRef.ForMovingXmlNodes();
            OfficeXmlNode root = part.Root;
            if (!root.Is(TransitionPresentationNamespace, "sld")) {
                throw new InvalidDataException("The slide part has no slide root.");
            }
            string[] removedIds = GetTransitionSoundRelationshipIds();
            string relationshipId = !part.IsSdk && part.Document is { } document
                && document.TryAddPackageTransitionSound(part.Name, audio, contentType, extension, out string packageId)
                    ? packageId : PowerPointEmbeddedSound.Add(SlidePart, audio, contentType, extension);

            OfficeXmlNode createdTransition = default;
            var added = new List<(OfficeXmlNode Transition, OfficeXmlNode Sound)>();
            try {
                // The audio writer can open the SDK for a compatibility case. Use the root now holding the slide.
                part = part.Current;
                root = part.Root;
                List<OfficeXmlNode> transitions = GetTransitionNodes(root).ToList();
                if (transitions.Count == 0) {
                    createdTransition = root.GetOrAddChild(TransitionPresentationNamespace, "transition", PSlideView.Order);
                    transitions.Add(createdTransition);
                }

                foreach (OfficeXmlNode transition in transitions) {
                    OfficeXmlNode action = transition.CreateElement(TransitionPresentationNamespace, "sndAc");
                    OfficeXmlNode start = action.Append(action.CreateElement(TransitionPresentationNamespace, "stSnd"));
                    start.SetBooleanAttribute("loop", loop);
                    OfficeXmlNode sound = start.Append(start.CreateElement(TransitionPresentationNamespace, "snd"));
                    sound.SetAttribute(OfficeXmlNamespaces.Relationships, "embed", relationshipId);
                    sound.SetAttribute("name", name);
                    added.Add((transition, transition.Append(action)));
                }
            } catch {
                foreach (var (_, sound) in added) sound.Remove();
                createdTransition.Remove();
                part.RemoveRelationshipIfUnused(relationshipId, part);
                throw;
            }

            // Keep all the previous actions until every branch has accepted its replacement.
            foreach ((OfficeXmlNode transition, OfficeXmlNode sound) in added) {
                foreach (OfficeXmlNode previous in transition.Children(TransitionPresentationNamespace, "sndAc").ToList()) {
                    if (previous != sound) previous.Remove();
                }
            }
            part.MarkEdited();
            RemoveUnusedTransitionSounds(removedIds);
        }

        /// <summary>Stops any currently playing transition sound on entry to this slide.</summary>
        public void StopTransitionSound() {
            PowerPointPartRef part = PartRef.ForMovingXmlNodes();
            string[] removedIds = GetTransitionSoundRelationshipIds();
            foreach (OfficeXmlNode transition in GetOrCreateTransitionNodes(part.Root)) {
                OfficeXmlNode action = ReplaceSoundAction(transition);
                action.Append(action.CreateElement(TransitionPresentationNamespace, "endSnd"));
            }
            part.MarkEdited();
            RemoveUnusedTransitionSounds(removedIds);
        }

        /// <summary>Removes this slide's start- or stop-sound transition action.</summary>
        public void ClearTransitionSound() {
            PowerPointPartRef part = PartRef.ForMovingXmlNodes();
            OfficeXmlNode root = part.Root;
            OfficeXmlNode direct = root.Child(TransitionPresentationNamespace, "transition");
            string[] removedIds = GetTransitionSoundRelationshipIds();
            foreach (OfficeXmlNode transition in GetTransitionNodes(root).ToList()) {
                transition.RemoveChildren(TransitionPresentationNamespace, "sndAc");
                if (transition == direct && transition.ChildNodes().Count == 0 && !transition.HasAttributes) {
                    transition.Remove();
                }
                part.MarkEdited();
            }
            RemoveUnusedTransitionSounds(removedIds);
        }

        /// <summary>Returns the exact embedded transition sound bytes, when present.</summary>
        public byte[]? GetTransitionSoundBytes() =>
            ReadSlideSound(GetTransitionSoundRelationshipId(GetTransitionNode()));

        private byte[]? ReadSlideSound(string? id) {
            PowerPointPartRef part = PartRef;
            if (!TryGetSlideAudioRelationship(part, id, out PowerPointRelationship relationship)) return null;
            PowerPointPartRef target = relationship.TargetPart;
            if (!target.IsNone) {
                if (target.Document?.ReadablePackage is { } package && package.TryGetStoredLength(target.Name, out long length)
                    && length > PowerPointMedia.MaximumOwnedPayloadBytes) {
                    return PowerPointEmbeddedSound.Read(SlidePart, id);
                }
                return target.ReadBytes();
            }
            return part.IsSdk ? PowerPointEmbeddedSound.Read(SlidePart, id) : null;
        }

        private static OfficeXmlNode TransitionStartSound(OfficeXmlNode transition) =>
            transition.Child(TransitionPresentationNamespace, "sndAc").Child(TransitionPresentationNamespace, "stSnd");

        private static OfficeXmlNode TransitionSound(OfficeXmlNode transition) =>
            TransitionStartSound(transition).Child(TransitionPresentationNamespace, "snd");

        private static string? GetTransitionSoundRelationshipId(OfficeXmlNode transition) =>
            TransitionSound(transition).Attribute(OfficeXmlNamespaces.Relationships, "embed");

        private string[] GetTransitionSoundRelationshipIds() =>
            GetTransitionNodes(PartRef.Root).Select(GetTransitionSoundRelationshipId)
                .Where(id => !string.IsNullOrEmpty(id)).Cast<string>().Distinct(StringComparer.Ordinal).ToArray();

        private void RemoveUnusedTransitionSounds(IEnumerable<string> relationshipIds) {
            PowerPointPartRef part = PartRef;
            foreach (string id in relationshipIds) {
                if (TryGetSlideAudioRelationship(part, id, out _)) part.RemoveRelationshipIfUnused(id, part);
            }
        }

        private static bool TryGetSlideAudioRelationship(PowerPointPartRef part, string? id, out PowerPointRelationship relationship) {
            relationship = default;
            return id != null && part.TryGetRelationship(id, out relationship) && !relationship.IsExternal &&
                OfficeXmlNamespaces.NormalizeRelationshipType(relationship.Type) == TransitionAudioRelationship;
        }

        private static OfficeXmlNode ReplaceSoundAction(OfficeXmlNode transition) {
            transition.RemoveChildren(TransitionPresentationNamespace, "sndAc");
            return transition.Append(transition.CreateElement(TransitionPresentationNamespace, "sndAc"));
        }

        private static IReadOnlyList<OfficeXmlNode> GetOrCreateTransitionNodes(OfficeXmlNode root) {
            List<OfficeXmlNode> existing = GetTransitionNodes(root).ToList();
            if (existing.Count != 0) return existing;
            return new[] { root.GetOrAddChild(TransitionPresentationNamespace, "transition", PSlideView.Order) };
        }
    }
}
