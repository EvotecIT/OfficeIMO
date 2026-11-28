using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        internal sealed class LegacyPptSoundProjectionContext {
            private readonly LegacyPptDocument _presentation;
            private readonly Dictionary<uint, LegacyPptImportPart> _mediaParts = new();
            private readonly Dictionary<(string Owner, uint SoundId), string> _relationships = new();

            internal LegacyPptSoundProjectionContext(
                LegacyPptDocument presentation) {
                _presentation = presentation
                    ?? throw new ArgumentNullException(nameof(presentation));
            }

            /// <summary>
            /// Relates <paramref name="ownerPart"/> to the embedded sound, adding the media part the first time any part
            /// uses it, and returns the relationship id; false when the binary sound has no projectable audio.
            /// </summary>
            internal bool TryProject(PowerPointPartRef ownerPart, uint soundId,
                out LegacyPptSound? sound, out string? relationshipId) {
                sound = _presentation.FindSound(soundId);
                relationshipId = null;
                if (sound?.ContentType == null || !sound.HasData) return false;
                PowerPointDocument document = ownerPart.Document
                    ?? throw new InvalidOperationException(
                        "Unable to resolve the target package for legacy sound import.");
                var key = (ownerPart.Name, soundId);
                if (_relationships.TryGetValue(key, out relationshipId)) return true;

                LegacyPptImportParts parts = LegacyPptImportParts.For(document);
                if (_mediaParts.TryGetValue(soundId, out LegacyPptImportPart mediaPart)) {
                    relationshipId = parts.AddRelationship(ownerPart, LegacyAudioRelationshipType,
                        mediaPart, generatedId: false);
                } else {
                    string extension = sound.ContentType == "audio/aiff" ? ".aiff" : ".wav";
                    relationshipId = parts.AddBinaryPart(ownerPart, LegacyAudioRelationshipType,
                        "/ppt/media/media" + extension, sound.ContentType, sound.DataBytes,
                        numbered: true, generatedId: false, out mediaPart);
                    _mediaParts.Add(soundId, mediaPart);
                }
                _relationships.Add(key, relationshipId);
                return true;
            }
        }
    }
}
