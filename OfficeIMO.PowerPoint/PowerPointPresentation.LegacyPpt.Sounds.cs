using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        internal sealed class LegacyPptSoundProjectionContext {
            private readonly LegacyPptPresentation _presentation;
            private readonly Dictionary<uint, MediaDataPart> _mediaParts = new();
            private readonly Dictionary<(OpenXmlPart Owner, uint SoundId), string>
                _relationships = new();

            internal LegacyPptSoundProjectionContext(
                LegacyPptPresentation presentation) {
                _presentation = presentation
                    ?? throw new ArgumentNullException(nameof(presentation));
            }

            internal bool TryProject(OpenXmlPart ownerPart, uint soundId,
                out LegacyPptSound? sound, out string? relationshipId) {
                sound = _presentation.FindSound(soundId);
                relationshipId = null;
                if (sound?.ContentType == null || !sound.HasData) return false;
                var key = (ownerPart, soundId);
                if (_relationships.TryGetValue(key, out relationshipId)) return true;

                if (!_mediaParts.TryGetValue(soundId, out MediaDataPart? mediaPart)) {
                    OpenXmlPackage package = GetPackage(ownerPart)
                        ?? throw new InvalidOperationException(
                            "Unable to resolve the target package for legacy sound import.");
                    mediaPart = CreateMediaDataPart(package, sound.ContentType);
                    using var input = new MemoryStream(sound.DataBytes, writable: false);
                    mediaPart.FeedData(input);
                    _mediaParts.Add(soundId, mediaPart);
                }

                relationshipId = GetNextRelationshipId(ownerPart);
                if (!TryAddAudioReferenceRelationship(ownerPart, mediaPart,
                        relationshipId)) {
                    throw new InvalidOperationException(
                        $"Unable to attach binary sound {soundId} to {ownerPart.GetType().Name}.");
                }
                _relationships.Add(key, relationshipId);
                return true;
            }
        }
    }
}
