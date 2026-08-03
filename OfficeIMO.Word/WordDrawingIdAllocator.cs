using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Runtime.CompilerServices;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Word {
    /// <summary>Allocates document-scoped non-visual DrawingML identifiers.</summary>
    internal static class WordDrawingIdAllocator {
        private static readonly ConditionalWeakTable<WordprocessingDocument, ReservationState> _reservations = new();

        internal static uint Reserve(WordDocument document, uint count = 1U) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return Reserve(document._wordprocessingDocument, count);
        }

        internal static uint Reserve(WordprocessingDocument document, uint count = 1U) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (count == 0U) throw new ArgumentOutOfRangeException(nameof(count));

            ReservationState state = _reservations.GetOrCreateValue(document);
            lock (state.SyncRoot) {
                uint max = Math.Max(FindPackageMaximum(document.MainDocumentPart), state.HighestReserved);
                if (max > uint.MaxValue - count) {
                    throw new InvalidOperationException("The document has exhausted the available DrawingML identifier range.");
                }

                uint first = max + 1U;
                state.HighestReserved = max + count;
                return first;
            }
        }

        internal static uint Reserve(MainDocumentPart mainPart, uint count = 1U) {
            if (mainPart == null) throw new ArgumentNullException(nameof(mainPart));
            if (mainPart.OpenXmlPackage is not WordprocessingDocument document) {
                throw new InvalidOperationException("The main document part is not attached to a Wordprocessing document.");
            }
            return Reserve(document, count);
        }

        internal static void Reassign(WordDocument document, OpenXmlElement root) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            Reassign(document._wordprocessingDocument, root);
        }

        internal static void Reassign(WordprocessingDocument document, OpenXmlElement root) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (root == null) throw new ArgumentNullException(nameof(root));

            OpenXmlElement[] elements = root.Descendants().Where(IsTracked).ToArray();
            if (elements.Length == 0) return;

            uint next = Reserve(document, checked((uint)elements.Length));
            foreach (OpenXmlElement element in elements) {
                SetId(element, next++);
            }
        }

        private static uint FindPackageMaximum(MainDocumentPart? mainPart) {
            uint max = 0U;
            UpdateMax(mainPart?.Document, ref max);
            if (mainPart == null) return max;

            foreach (HeaderPart headerPart in mainPart.HeaderParts) UpdateMax(headerPart.Header, ref max);
            foreach (FooterPart footerPart in mainPart.FooterParts) UpdateMax(footerPart.Footer, ref max);
            UpdateMax(mainPart.FootnotesPart?.Footnotes, ref max);
            UpdateMax(mainPart.EndnotesPart?.Endnotes, ref max);
            UpdateMax(mainPart.WordprocessingCommentsPart?.Comments, ref max);
            return max;
        }

        private static void UpdateMax(OpenXmlElement? root, ref uint max) {
            if (root == null) return;

            foreach (OpenXmlElement element in root.Descendants()) {
                uint? id = element switch {
                    DW.DocProperties properties => properties.Id?.Value,
                    PIC.NonVisualDrawingProperties properties => properties.Id?.Value,
                    Wpg.NonVisualDrawingProperties properties => properties.Id?.Value,
                    Wps.NonVisualDrawingProperties properties => properties.Id?.Value,
                    _ => null,
                };
                if (id.HasValue && id.Value > max) max = id.Value;
            }
        }

        private static bool IsTracked(OpenXmlElement element) =>
            element is DW.DocProperties or PIC.NonVisualDrawingProperties or
                Wpg.NonVisualDrawingProperties or Wps.NonVisualDrawingProperties;

        private static void SetId(OpenXmlElement element, uint id) {
            switch (element) {
                case DW.DocProperties properties: properties.Id = id; break;
                case PIC.NonVisualDrawingProperties properties: properties.Id = id; break;
                case Wpg.NonVisualDrawingProperties properties: properties.Id = id; break;
                case Wps.NonVisualDrawingProperties properties: properties.Id = id; break;
                default: throw new ArgumentOutOfRangeException(nameof(element));
            }
        }

        private sealed class ReservationState {
            internal object SyncRoot { get; } = new();
            internal uint HighestReserved { get; set; }
        }
    }
}
