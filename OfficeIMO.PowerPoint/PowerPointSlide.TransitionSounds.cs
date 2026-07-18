using DocumentFormat.OpenXml.Presentation;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>Gets whether this slide starts an embedded transition sound.</summary>
        public bool HasTransitionSound => GetTransitionStartSound()?.Sound != null;

        /// <summary>Gets the embedded transition sound name, when present.</summary>
        public string? TransitionSoundName =>
            GetTransitionStartSound()?.Sound?.Name?.Value;

        /// <summary>Gets whether the embedded transition sound loops.</summary>
        public bool TransitionSoundLoops =>
            GetTransitionStartSound()?.Loop?.Value == true;

        /// <summary>Gets whether this transition stops the currently playing sound.</summary>
        public bool TransitionStopsSound => GetTransitionElement()?
            .GetFirstChild<SoundAction>()?
            .GetFirstChild<EndSoundAction>() != null;

        /// <summary>Sets an embedded WAV or AIFF transition sound from a file.</summary>
        public void SetTransitionSound(string audioPath, bool loop = false) {
            if (audioPath == null) throw new ArgumentNullException(nameof(audioPath));
            if (!File.Exists(audioPath)) {
                throw new FileNotFoundException("Audio file not found.", audioPath);
            }
            using FileStream input = new(audioPath, FileMode.Open, FileAccess.Read,
                FileShare.Read);
            SetTransitionSound(input, Path.GetFileName(audioPath),
                GetAudioContentType(audioPath), Path.GetExtension(audioPath), loop);
        }

        /// <summary>Sets an embedded WAV or AIFF transition sound.</summary>
        public void SetTransitionSound(Stream audio, string name,
            string contentType = "audio/wav", string extension = ".wav",
            bool loop = false) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("A transition sound name is required.",
                    nameof(name));
            }
            string[] previousRelationshipIds =
                GetTransitionSoundRelationshipIds();
            string relationshipId = PowerPointEmbeddedSound.Add(_slidePart,
                audio, contentType, extension);
            foreach (Transition transition in
                     GetOrCreateTransitionElements()) {
                transition.RemoveAllChildren<SoundAction>();
                transition.Append(new SoundAction(
                    new StartSoundAction(new Sound {
                        Embed = relationshipId,
                        Name = name
                    }) { Loop = loop }));
            }
            RemoveUnusedTransitionSounds(previousRelationshipIds);
        }

        /// <summary>Stops any currently playing transition sound on entry to this slide.</summary>
        public void StopTransitionSound() {
            string[] previousRelationshipIds =
                GetTransitionSoundRelationshipIds();
            foreach (Transition transition in
                     GetOrCreateTransitionElements()) {
                transition.RemoveAllChildren<SoundAction>();
                transition.Append(new SoundAction(new EndSoundAction()));
            }
            RemoveUnusedTransitionSounds(previousRelationshipIds);
        }

        /// <summary>Removes this slide's start- or stop-sound transition action.</summary>
        public void ClearTransitionSound() {
            string[] relationshipIds =
                GetTransitionSoundRelationshipIds();
            foreach (Transition transition in GetTransitionElements()) {
                transition.RemoveAllChildren<SoundAction>();
                if (ReferenceEquals(transition, SlideRoot.Transition)
                    && transition.ChildElements.Count == 0
                    && !transition.HasAttributes) {
                    SlideRoot.Transition = null;
                }
            }
            RemoveUnusedTransitionSounds(relationshipIds);
        }

        /// <summary>Returns the exact embedded transition sound bytes, when present.</summary>
        public byte[]? GetTransitionSoundBytes() {
            string? relationshipId = GetTransitionStartSound()?.Sound?.Embed?.Value;
            return PowerPointEmbeddedSound.Read(_slidePart, relationshipId);
        }

        private StartSoundAction? GetTransitionStartSound() =>
            GetTransitionElement()?.GetFirstChild<SoundAction>()?
                .GetFirstChild<StartSoundAction>();

        private string[] GetTransitionSoundRelationshipIds() =>
            GetTransitionElements()
                .Select(transition => transition
                    .GetFirstChild<SoundAction>()?
                    .GetFirstChild<StartSoundAction>()?.Sound?.Embed?.Value)
                .Where(id => !string.IsNullOrEmpty(id))
                .Cast<string>()
                .Distinct(StringComparer.Ordinal)
                .ToArray();

        private IReadOnlyList<Transition> GetOrCreateTransitionElements() {
            IReadOnlyList<Transition> transitions =
                GetTransitionElements();
            if (transitions.Count > 0) return transitions;
            return new[] { SlideRoot.Transition = new Transition() };
        }
    }
}
