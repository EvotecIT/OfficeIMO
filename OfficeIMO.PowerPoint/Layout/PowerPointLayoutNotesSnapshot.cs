using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    /// <summary>Restores the relationship targets as well as the notes XML after rejected collected content.</summary>
    internal sealed class PowerPointLayoutNotesSnapshot {
        private readonly NotesSlidePart _part;
        private readonly IdPartPair[] _parts;
        private readonly (string Id, Uri Uri, bool External)[] _links;
        private readonly (string Id, string ContentType, string Extension, byte[] Bytes)[] _audio;

        internal PowerPointLayoutNotesSnapshot(NotesSlidePart part) {
            _part = part;
            _parts = part.Parts.ToArray();
            _links = part.HyperlinkRelationships.Select(r => (r.Id, r.Uri, r.IsExternal)).ToArray();
            _audio = part.DataPartReferenceRelationships.OfType<AudioReferenceRelationship>().Select(r => {
                using var input = r.DataPart.GetStream(FileMode.Open, FileAccess.Read);
                using var output = new MemoryStream();
                input.CopyTo(output);
                return (r.Id, r.DataPart.ContentType, Path.GetExtension(r.DataPart.Uri.OriginalString), output.ToArray());
            }).ToArray();
        }

        internal void Restore() {
            foreach (var relationship in _part.Parts.ToArray()) {
                if (!_parts.Any(p => p.RelationshipId == relationship.RelationshipId && ReferenceEquals(p.OpenXmlPart, relationship.OpenXmlPart)))
                    _part.DeletePart(relationship.RelationshipId);
            }
            foreach (var relationship in _parts) {
                if (!_part.Parts.Any(p => p.RelationshipId == relationship.RelationshipId))
                    _part.AddPart(relationship.OpenXmlPart, relationship.RelationshipId);
            }
            foreach (var link in _part.HyperlinkRelationships.ToArray()) _part.DeleteReferenceRelationship(link);
            foreach (var link in _links) _part.AddHyperlinkRelationship(link.Uri, link.External, link.Id);

            var document = (PresentationDocument)_part.OpenXmlPackage;
            var current = _part.DataPartReferenceRelationships.OfType<AudioReferenceRelationship>().ToArray();
            var payloads = current.Select(r => r.DataPart).OfType<MediaDataPart>().ToArray();
            foreach (var relationship in current) _part.DeleteReferenceRelationship(relationship);
            PowerPointEmbeddedSound.RemoveUnreferencedMediaDataParts(document, payloads);
            foreach (var audio in _audio) {
                var media = document.CreateMediaDataPart(audio.ContentType, audio.Extension);
                using var stream = new MemoryStream(audio.Bytes, writable: false);
                media.FeedData(stream);
                _part.AddAudioReferenceRelationship(media, audio.Id);
            }
        }
    }
}
