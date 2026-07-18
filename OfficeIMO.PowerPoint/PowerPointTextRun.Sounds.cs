using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTextRun {
        /// <summary>Gets the sound name played when this text run is clicked.</summary>
        public string? ClickSoundName => GetInteractionSound(mouseOver: false)?
            .Name?.Value;

        /// <summary>Gets the sound name played when the pointer enters this text run.</summary>
        public string? MouseOverSoundName => GetInteractionSound(mouseOver: true)?
            .Name?.Value;

        /// <summary>Sets an embedded WAV or AIFF sound played when this text run is clicked.</summary>
        public void SetClickSound(Stream audio, string name,
            string contentType = "audio/wav", string extension = ".wav") =>
            SetInteractionSound(mouseOver: false, audio, name, contentType,
                extension);

        /// <summary>Sets an embedded WAV or AIFF sound played when the pointer enters this text run.</summary>
        public void SetMouseOverSound(Stream audio, string name,
            string contentType = "audio/wav", string extension = ".wav") =>
            SetInteractionSound(mouseOver: true, audio, name, contentType,
                extension);

        /// <summary>Controls whether clicking this text run stops the currently playing sound.</summary>
        public void SetClickStopsSound(bool value) =>
            SetInteractionStopsSound(mouseOver: false, value);

        /// <summary>Controls whether entering this text run stops the currently playing sound.</summary>
        public void SetMouseOverStopsSound(bool value) =>
            SetInteractionStopsSound(mouseOver: true, value);

        /// <summary>Removes the sound played when this text run is clicked.</summary>
        public void ClearClickSound() => ClearInteractionSound(mouseOver: false);

        /// <summary>Removes the sound played when the pointer enters this text run.</summary>
        public void ClearMouseOverSound() => ClearInteractionSound(mouseOver: true);

        /// <summary>Returns the exact embedded click-sound bytes, when present.</summary>
        public byte[]? GetClickSoundBytes() => GetInteractionSoundBytes(
            mouseOver: false);

        /// <summary>Returns the exact embedded mouse-over-sound bytes, when present.</summary>
        public byte[]? GetMouseOverSoundBytes() => GetInteractionSoundBytes(
            mouseOver: true);

        private void SetInteractionSound(bool mouseOver, Stream audio,
            string name, string contentType, string extension) {
            if (_slidePart == null) {
                throw new InvalidOperationException(
                    "Action sounds require a text run attached to a slide.");
            }
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("An action sound name is required.",
                    nameof(name));
            }
            string? previousRelationshipId = GetInteractionSound(mouseOver)?
                .Embed?.Value;
            string relationshipId = PowerPointEmbeddedSound.Add(_slidePart,
                audio, contentType, extension);
            A.HyperlinkType hyperlink = GetOrCreateInteraction(mouseOver);
            hyperlink.RemoveAllChildren<A.HyperlinkSound>();
            hyperlink.Append(new A.HyperlinkSound {
                Embed = relationshipId,
                Name = name
            });
            PowerPointEmbeddedSound.RemoveIfUnused(_slidePart,
                previousRelationshipId);
        }

        private void SetInteractionStopsSound(bool mouseOver, bool value) {
            A.HyperlinkType? hyperlink = value
                ? GetOrCreateInteraction(mouseOver)
                : GetInteraction(mouseOver);
            if (hyperlink != null) hyperlink.EndSound = value ? true : null;
        }

        private void ClearInteractionSound(bool mouseOver) {
            if (_slidePart == null) return;
            A.HyperlinkType? hyperlink = GetInteraction(mouseOver);
            string? relationshipId = hyperlink?
                .GetFirstChild<A.HyperlinkSound>()?.Embed?.Value;
            hyperlink?.RemoveAllChildren<A.HyperlinkSound>();
            PowerPointEmbeddedSound.RemoveIfUnused(_slidePart,
                relationshipId);
        }

        private byte[]? GetInteractionSoundBytes(bool mouseOver) {
            if (_slidePart == null) return null;
            return PowerPointEmbeddedSound.Read(_slidePart,
                GetInteractionSound(mouseOver)?.Embed?.Value);
        }

        private A.HyperlinkSound? GetInteractionSound(bool mouseOver) =>
            GetInteraction(mouseOver)?.GetFirstChild<A.HyperlinkSound>();

        private A.HyperlinkType? GetInteraction(bool mouseOver) {
            A.RunProperties? properties = Run.RunProperties;
            return mouseOver
                ? properties?.GetFirstChild<A.HyperlinkOnMouseOver>()
                : properties?.GetFirstChild<A.HyperlinkOnClick>();
        }

        private A.HyperlinkType GetOrCreateInteraction(bool mouseOver) {
            A.RunProperties properties = EnsureRunProperties();
            A.HyperlinkType? existing = mouseOver
                ? properties.GetFirstChild<A.HyperlinkOnMouseOver>()
                : properties.GetFirstChild<A.HyperlinkOnClick>();
            if (existing != null) return existing;
            A.HyperlinkType created = mouseOver
                ? new A.HyperlinkOnMouseOver { Id = string.Empty }
                : new A.HyperlinkOnClick { Id = string.Empty };
            properties.Append(created);
            return created;
        }
    }
}
