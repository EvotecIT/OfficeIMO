using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// The relationship side of layout materialization: moving content between parts together with the relationships it
    /// uses, releasing relationships materialization no longer uses, and restoring notes after rejected content.
    /// </summary>
    public partial class PowerPointSlide {
        /// <summary>Transfers a layout-owned text frame without replacing its native wrapper or body.</summary>
        internal void AdoptLayoutTextBox(PowerPointTextBox box) {
            AdoptLayoutShape(box);
        }

        /// <summary>Moves a materialized shape together with its interaction relationships.</summary>
        internal void AdoptLayoutShape(PowerPointShape box) {
            if (ReferenceEquals(box.OwnerSlide, this)) return;
            var drawing = box.Element.Descendants<NonVisualDrawingProperties>().First();
            foreach (var child in box.Element.ChildElements) {
                // Paragraphs and table rows bind individually: detached later slices still own source IDs.
                if (child is TextBody || box is PowerPointTable && child is A.Graphic) continue;
                RehomeLayoutRelationships(child, box.RelationshipOwnerPart, SlidePart);
            }
            box.OwnerSlide?.ShapeList.Remove(box);
            if (box.Element.Parent != null) box.Element.Remove();
            drawing.Id = AllocateShapeId();
            drawing.Name = GenerateUniqueName(box.Name ?? "Content");
            SlideRoot.CommonSlideData!.ShapeTree!.Append(box.Element);
            box.AttachTo(this, SlidePart);
            ShapeList.Add(box);
        }

        /// <summary>
        /// Gives <paramref name="element"/>, which moves to this slide from <paramref name="source"/>, relationships of this
        /// slide to the same targets, and points its relationship ids at them. Content already on this slide is left as it is.
        /// </summary>
        internal void RehomeLayoutRelationships(OfficeXmlNode element, PowerPointPartRef source) {
            if (source.IsNone) return;
            PowerPointPartRef current = source.Current;
            if (current.SdkPart == null && current.Document is { } document) {
                _ = document.OpenXmlDocument;
                current = source.Current;
            }
            if (current.SdkPart is not OpenXmlPart part || ReferenceEquals(part, SlidePart)) return;
            OpenXmlElement target = element.SdkElement as OpenXmlElement
                ?? throw new InvalidOperationException("The content to move is not held in an SDK part.");
            RehomeLayoutRelationships(target, part, SlidePart);
        }

        internal static void RehomeLayoutRelationships(OpenXmlElement element, OpenXmlPart? source, SlidePart target) {
            if (source == null || ReferenceEquals(source, target)) return;
            var ids = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var node in new[] { element }.Concat(element.Descendants())) {
                foreach (var attribute in node.GetAttributes().ToArray()) {
                    if (attribute.NamespaceUri != PowerPointUtils.RelationshipIdNamespace || string.IsNullOrEmpty(attribute.Value)) continue;
                    string id = attribute.Value!;
                    if (!ids.TryGetValue(id, out string? replacement)) {
                        var audio = PowerPointEmbeddedSound.Find(source, id);
                        var hyperlink = source.HyperlinkRelationships.FirstOrDefault(r => r.Id == id);
                        var part = source.Parts.FirstOrDefault(p => p.RelationshipId == id).OpenXmlPart;
                        var external = source.ExternalRelationships.FirstOrDefault(r => r.Id == id);
                        if (audio != null) replacement = target.AddAudioReferenceRelationship(audio).Id;
                        else if (hyperlink != null) replacement = target.HyperlinkRelationships
                            .FirstOrDefault(candidate => candidate.Uri == hyperlink.Uri && candidate.IsExternal == hyperlink.IsExternal)?.Id
                            ?? target.AddHyperlinkRelationship(hyperlink.Uri, hyperlink.IsExternal).Id;
                        else if (part != null) {
                            target.AddPart(part);
                            replacement = target.GetIdOfPart(part);
                        } else if (external != null) replacement = target.AddExternalRelationship(external.RelationshipType, external.Uri).Id;
                        else continue;
                        ids.Add(id, replacement);
                    }
                    node.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, replacement));
                }
            }
        }

        /// <summary>Release source references only after every detached continuation slice has bound.</summary>
        internal void RemoveUnusedMaterializationRelationships() {
            SlidePart part = SlidePart;
            var used = new HashSet<string>(part.Slide!.Descendants().SelectMany(node => node.GetAttributes())
                .Where(attribute => attribute.NamespaceUri == PowerPointUtils.RelationshipIdNamespace)
                .Select(attribute => attribute.Value).OfType<string>(), StringComparer.Ordinal);
            foreach (var link in part.HyperlinkRelationships.Where(link => !used.Contains(link.Id)).ToArray())
                part.DeleteReferenceRelationship(link);
            var audio = part.DataPartReferenceRelationships.OfType<AudioReferenceRelationship>()
                .Where(link => !used.Contains(link.Id)).ToArray();
            var media = audio.Select(link => link.DataPart).OfType<MediaDataPart>().ToArray();
            foreach (var link in audio) part.DeleteReferenceRelationship(link);
            PowerPointEmbeddedSound.RemoveUnreferencedMediaDataParts((PresentationDocument)part.OpenXmlPackage, media);
            foreach (var child in part.Parts.Where(child => child.OpenXmlPart is ImagePart or ChartPart
                         && !used.Contains(child.RelationshipId)).ToArray())
                part.DeletePart(child.RelationshipId);
        }

        /// <summary>Captures notes structurally so a failed collected column cannot leave text or empty note parts.</summary>
        internal Action CaptureLayoutNotesRollback() {
            var part = SlidePart;
            var presentation = part.GetParentParts().OfType<PresentationPart>().Single();
            var notes = part.NotesSlidePart;
            var root = (NotesSlide?)notes?.NotesSlide?.CloneNode(true);
            var relationships = notes == null ? null : new LayoutNotesSnapshot(notes);
            var master = presentation.NotesMasterPart;
            var presentationRoot = presentation.Presentation!;
            var masterIds = (NotesMasterIdList?)presentationRoot.NotesMasterIdList?.CloneNode(true);
            return () => {
                if (notes == null) {
                    if (part.NotesSlidePart != null) {
                        var media = PowerPointEmbeddedSound.GetReferencedMediaDataParts(part.NotesSlidePart);
                        part.DeletePart(part.NotesSlidePart);
                        PowerPointEmbeddedSound.RemoveUnreferencedMediaDataParts((PresentationDocument)part.OpenXmlPackage, media);
                    }
                } else {
                    relationships!.Restore();
                    notes.NotesSlide = root;
                }
                if (master == null && presentation.NotesMasterPart != null && !presentation.SlideParts.Any(s => s.NotesSlidePart != null)) {
                    presentation.DeletePart(presentation.NotesMasterPart);
                    presentationRoot.NotesMasterIdList = masterIds;
                }
            };
        }

        /// <summary>Restores the relationship targets as well as the notes XML after rejected collected content.</summary>
        private sealed class LayoutNotesSnapshot {
            private readonly NotesSlidePart _part;
            private readonly IdPartPair[] _parts;
            private readonly (string Id, Uri Uri, bool External)[] _links;
            private readonly (string Id, string ContentType, string Extension, byte[] Bytes)[] _audio;

            internal LayoutNotesSnapshot(NotesSlidePart part) {
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
}
