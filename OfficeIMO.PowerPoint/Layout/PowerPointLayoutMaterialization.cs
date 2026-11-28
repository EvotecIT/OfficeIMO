using System.Runtime.CompilerServices;
using Ns = OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces;

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
            box.XmlElement.Child(Ns.PresentationML, "txBody").RemoveChildren(Ns.DrawingML, "p");
            box.OwnerPartRef.MarkEdited();
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
            foreach (var slide in slides) slide.RemoveUnusedMaterializationRelationships();
            _completed = true;
            return slides;
        }

        /// <summary>Captures notes structurally so a failed collected column cannot leave text or empty note parts.</summary>
        internal static Action CaptureNotesRollback(PowerPointSlide slide) => slide.CaptureLayoutNotesRollback();

        public void Dispose() {
            if (_completed) return;
            _rollbackNotes();
            while (_document.Slides.Count > _start) _document.RemoveSlide(_document.Slides.Count - 1);
            _rollback();
        }
    }
}
