using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

[assembly: InternalsVisibleTo("PSWriteOffice")]

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Owns the native slide used while a content block executes. Completion arranges its live bodies through
    /// the normal layout engine; disposal before completion removes only this operation's appended slides.
    /// </summary>
    internal sealed class PowerPointLayoutMaterialization : IDisposable {
        private readonly PowerPointDocument _document;
        private readonly int _start;
        private readonly Action _rollback;
        private readonly Action _rollbackNotes;
        private bool _completed;

        internal PowerPointLayoutMaterialization(PowerPointDocument document) {
            _document = document;
            _start = document.Slides.Count;
            _rollback = document.LayoutDeck.CaptureRollback();
            Slide = document.AddSlide();
            _rollbackNotes = CaptureNotesRollback(Slide);
        }

        internal PowerPointSlide Slide { get; }

        internal PowerPointTextBox CreateTextBody() {
            var box = Slide.AddTextBox(string.Empty);
            box.TextBody!.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.Paragraph>();
            return box;
        }

        internal IReadOnlyList<PowerPointSlide> Complete(PowerPointContentPlanSlide plan, PowerPointCompositionOptions options,
            Action<IReadOnlyList<PowerPointSlide>>? finalize = null) {
            plan.MaterializedSlide = Slide;
            // Notes may already have been edited through the returned native slide.
            plan.Notes = null;
            var slides = _document.Compose(new PowerPointDeckPlan().Add(plan), options).Slides;
            // Appearance applies to every physical continuation before the transaction is accepted.
            finalize?.Invoke(slides);
            foreach (var slide in slides) RemoveUnusedMaterializationRelationships(slide.SlidePart);
            _completed = true;
            return slides;
        }

        /// <summary>Release source references only after every detached continuation slice has bound.</summary>
        private static void RemoveUnusedMaterializationRelationships(SlidePart part) {
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
        internal static Action CaptureNotesRollback(PowerPointSlide slide) {
            var part = slide.SlidePart;
            var presentation = part.GetParentParts().OfType<PresentationPart>().Single();
            var notes = part.NotesSlidePart;
            var root = (NotesSlide?)notes?.NotesSlide?.CloneNode(true);
            var relationships = notes == null ? null : new PowerPointLayoutNotesSnapshot(notes);
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

        public void Dispose() {
            if (_completed) return;
            _rollbackNotes();
            while (_document.Slides.Count > _start) _document.RemoveSlide(_document.Slides.Count - 1);
            _rollback();
        }
    }
}
