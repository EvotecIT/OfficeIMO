using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointEmbeddedSound {
        internal static string Add(SlidePart slidePart, Stream audio,
            string contentType, string extension) {
            if (slidePart == null) throw new ArgumentNullException(nameof(slidePart));
            if (audio == null) throw new ArgumentNullException(nameof(audio));
            if (!audio.CanRead) {
                throw new ArgumentException("Audio stream must be readable.",
                    nameof(audio));
            }
            if (string.IsNullOrWhiteSpace(contentType)) {
                throw new ArgumentException("An audio content type is required.",
                    nameof(contentType));
            }
            string normalizedContentType = contentType.Trim().ToLowerInvariant();
            if (normalizedContentType != "audio/wav"
                && normalizedContentType != "audio/x-wav"
                && normalizedContentType != "audio/aiff"
                && normalizedContentType != "audio/x-aiff") {
                throw new NotSupportedException(
                    "Transition, action, and animation sounds require embedded WAV or AIFF audio.");
            }
            string normalizedExtension = extension?.Trim().TrimStart('.')
                ?? string.Empty;
            if (normalizedExtension.Length == 0) {
                throw new ArgumentException("An audio file extension is required.",
                    nameof(extension));
            }
            if (slidePart.OpenXmlPackage is not PresentationDocument document) {
                throw new InvalidOperationException(
                    "The slide is not attached to a presentation document.");
            }
            MediaDataPart mediaPart = document.CreateMediaDataPart(contentType,
                normalizedExtension);
            if (audio.CanSeek) audio.Position = 0;
            mediaPart.FeedData(audio);
            string relationshipId = GetNextRelationshipId(slidePart);
            slidePart.AddAudioReferenceRelationship(mediaPart, relationshipId);
            return relationshipId;
        }

        internal static MediaDataPart? Find(SlidePart slidePart,
            string? relationshipId) {
            if (string.IsNullOrEmpty(relationshipId)) return null;
            return slidePart.DataPartReferenceRelationships
                .OfType<AudioReferenceRelationship>()
                .FirstOrDefault(relationship => string.Equals(relationship.Id,
                    relationshipId, StringComparison.Ordinal))?
                .DataPart as MediaDataPart;
        }

        internal static byte[]? Read(SlidePart slidePart,
            string? relationshipId) {
            MediaDataPart? mediaPart = Find(slidePart, relationshipId);
            if (mediaPart == null) return null;
            using Stream input = mediaPart.GetStream(FileMode.Open, FileAccess.Read);
            using var output = new MemoryStream();
            input.CopyTo(output);
            return output.ToArray();
        }

        internal static void RemoveIfUnused(SlidePart slidePart,
            string? relationshipId) {
            if (slidePart == null) {
                throw new ArgumentNullException(nameof(slidePart));
            }
            if (string.IsNullOrEmpty(relationshipId)
                || slidePart.Slide?.Descendants().Any(element =>
                    element.GetAttributes().Any(attribute => string.Equals(
                        attribute.Value, relationshipId,
                        StringComparison.Ordinal))) == true) {
                return;
            }
            AudioReferenceRelationship? relationship = slidePart
                .DataPartReferenceRelationships
                .OfType<AudioReferenceRelationship>()
                .FirstOrDefault(candidate => string.Equals(candidate.Id,
                    relationshipId, StringComparison.Ordinal));
            if (relationship == null) return;
            MediaDataPart? mediaPart = relationship.DataPart
                as MediaDataPart;
            slidePart.DeleteReferenceRelationship(relationship);
            if (mediaPart != null
                && !mediaPart.GetDataPartReferenceRelationships().Any()
                && slidePart.OpenXmlPackage is PresentationDocument document) {
                document.DeletePart(mediaPart);
            }
        }

        private static string GetNextRelationshipId(SlidePart slidePart) {
            var used = new HashSet<string>(slidePart.Parts.Select(pair =>
                    pair.RelationshipId)
                .Concat(slidePart.ExternalRelationships.Select(item => item.Id))
                .Concat(slidePart.HyperlinkRelationships.Select(item => item.Id))
                .Concat(slidePart.DataPartReferenceRelationships.Select(item =>
                    item.Id)), StringComparer.Ordinal);
            int next = 1;
            string value;
            do {
                value = "rId" + next++;
            } while (!used.Add(value));
            return value;
        }
    }
}
