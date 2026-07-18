using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

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
            string relationshipId = GetNextRelationshipId(slidePart);
            try {
                if (audio.CanSeek) audio.Position = 0;
                mediaPart.FeedData(audio);
                slidePart.AddAudioReferenceRelationship(mediaPart,
                    relationshipId);
                return relationshipId;
            } catch {
                AudioReferenceRelationship? relationship = slidePart
                    .DataPartReferenceRelationships
                    .OfType<AudioReferenceRelationship>()
                    .FirstOrDefault(candidate => ReferenceEquals(
                        candidate.DataPart, mediaPart));
                if (relationship != null) {
                    slidePart.DeleteReferenceRelationship(relationship);
                }
                if (!mediaPart.GetDataPartReferenceRelationships().Any()) {
                    document.DeletePart(mediaPart);
                }
                throw;
            }
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

        internal static void RemoveIfUnused(OpenXmlPart ownerPart,
            string? relationshipId) {
            if (ownerPart == null) {
                throw new ArgumentNullException(nameof(ownerPart));
            }
            string id = relationshipId ?? string.Empty;
            if (id.Length == 0) return;
            if (ReferencesRelationship(ownerPart.RootElement,
                    id)) return;
            AudioReferenceRelationship? relationship = ownerPart
                .DataPartReferenceRelationships
                .OfType<AudioReferenceRelationship>()
                .FirstOrDefault(candidate => string.Equals(candidate.Id,
                    id, StringComparison.Ordinal));
            if (relationship == null) return;
            MediaDataPart? mediaPart = relationship.DataPart
                as MediaDataPart;
            ownerPart.DeleteReferenceRelationship(relationship);
            if (mediaPart != null
                && !mediaPart.GetDataPartReferenceRelationships().Any()
                && ownerPart.OpenXmlPackage is PresentationDocument document) {
                document.DeletePart(mediaPart);
            }
        }

        internal static string[] GetRelationshipIds(
            DocumentFormat.OpenXml.OpenXmlElement container) {
            if (container == null) {
                throw new ArgumentNullException(nameof(container));
            }
            return container.Descendants<A.HyperlinkSound>()
                .Select(sound => sound.Embed?.Value)
                .Where(id => !string.IsNullOrEmpty(id))
                .Cast<string>()
                .Distinct(StringComparer.Ordinal)
                .ToArray();
        }

        internal static void RemoveIfUnused(OpenXmlPart? ownerPart,
            IEnumerable<string> relationshipIds) {
            if (ownerPart == null) return;
            if (relationshipIds == null) {
                throw new ArgumentNullException(nameof(relationshipIds));
            }
            foreach (string relationshipId in relationshipIds.Distinct(
                         StringComparer.Ordinal)) {
                RemoveIfUnused(ownerPart, relationshipId);
            }
        }

        /// <summary>
        /// Collects package-level media data referenced by an Open XML part
        /// subtree so callers can clean payloads whose owning parts are deleted.
        /// </summary>
        internal static MediaDataPart[] GetReferencedMediaDataParts(
            OpenXmlPart rootPart) {
            if (rootPart == null) {
                throw new ArgumentNullException(nameof(rootPart));
            }
            var mediaParts = new HashSet<MediaDataPart>();
            var visited = new HashSet<OpenXmlPart>();
            var pending = new Stack<OpenXmlPart>();
            pending.Push(rootPart);
            while (pending.Count > 0) {
                OpenXmlPart part = pending.Pop();
                if (!visited.Add(part)) continue;
                foreach (IdPartPair child in part.Parts) {
                    pending.Push(child.OpenXmlPart);
                }
                foreach (MediaDataPart mediaPart in part
                             .DataPartReferenceRelationships
                             .Select(relationship => relationship.DataPart)
                             .OfType<MediaDataPart>()) {
                    mediaParts.Add(mediaPart);
                }
            }
            return mediaParts.ToArray();
        }

        /// <summary>
        /// Deletes collected media payloads after their final package
        /// reference has been removed.
        /// </summary>
        internal static void RemoveUnreferencedMediaDataParts(
            PresentationDocument document,
            IEnumerable<MediaDataPart> mediaParts) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }
            if (mediaParts == null) {
                throw new ArgumentNullException(nameof(mediaParts));
            }
            foreach (MediaDataPart mediaPart in mediaParts.Distinct()) {
                if (mediaPart.GetDataPartReferenceRelationships().Any()) {
                    continue;
                }
                if (document.DataParts.Any(part =>
                        ReferenceEquals(part, mediaPart))) {
                    document.DeletePart(mediaPart);
                }
            }
        }

        private static bool ReferencesRelationship(
            OpenXmlPartRootElement? root, string relationshipId) =>
            root != null && (root.GetAttributes().Any(attribute =>
                string.Equals(attribute.NamespaceUri,
                    PowerPointUtils.RelationshipIdNamespace,
                    StringComparison.Ordinal)
                && string.Equals(attribute.Value, relationshipId,
                    StringComparison.Ordinal))
                || root.Descendants().Any(element => element.GetAttributes()
                    .Any(attribute => string.Equals(attribute.NamespaceUri,
                            PowerPointUtils.RelationshipIdNamespace,
                            StringComparison.Ordinal)
                        && string.Equals(attribute.Value, relationshipId,
                            StringComparison.Ordinal))));

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
