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
            ArgumentNullException.ThrowIfNull(document);
            return Reserve(document._wordprocessingDocument, count);
        }

        internal static uint Reserve(WordprocessingDocument document, uint count = 1U) {
            ArgumentNullException.ThrowIfNull(document);
            if (count == 0U) throw new ArgumentOutOfRangeException(nameof(count));

            ReservationState state = _reservations.GetOrCreateValue(document);
            lock (state.SyncRoot) {
                // The package is scanned once, so ids already present are never reused; later reservations continue
                // from the highest id handed out. Import paths invalidate the state to pick up ids they bring in.
                if (!state.Seeded || state.Shared) {
                    state.HighestReserved = Math.Max(FindPackageMaximum(document.MainDocumentPart), state.HighestReserved);
                    state.Seeded = true;
                }
                uint max = state.HighestReserved;
                if (max > uint.MaxValue - count) {
                    throw new InvalidOperationException("The document has exhausted the available DrawingML identifier range.");
                }

                uint first = max + 1U;
                state.HighestReserved = max + count;
                return first;
            }
        }

        internal static uint Reserve(MainDocumentPart mainPart, uint count = 1U) {
            ArgumentNullException.ThrowIfNull(mainPart);
            if (mainPart.OpenXmlPackage is not WordprocessingDocument document) {
                throw new InvalidOperationException("The main document part is not attached to a Wordprocessing document.");
            }
            return Reserve(document, count);
        }

        /// <summary>Makes the next reservation rescan the package for ids added outside the allocator.</summary>
        internal static void Invalidate(WordprocessingDocument document) {
            ArgumentNullException.ThrowIfNull(document);
            if (!_reservations.TryGetValue(document, out ReservationState? state)) return;
            lock (state.SyncRoot) state.Seeded = false;
        }

        /// <summary>
        /// Rescans on every reservation once the package has been handed to callers, whose direct edits can't be observed.
        /// </summary>
        internal static void MarkPackageShared(WordprocessingDocument document) {
            ArgumentNullException.ThrowIfNull(document);
            ReservationState state = _reservations.GetOrCreateValue(document);
            lock (state.SyncRoot) state.Shared = true;
        }

        internal static void Reassign(WordDocument document, OpenXmlElement root) {
            ArgumentNullException.ThrowIfNull(document);
            Reassign(document._wordprocessingDocument, root);
        }

        internal static void Reassign(WordprocessingDocument document, OpenXmlElement root) {
            ArgumentNullException.ThrowIfNull(document);
            ArgumentNullException.ThrowIfNull(root);

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

            foreach (HeaderPart headerPart in mainPart.HeaderParts) UpdateMax(headerPart, ref max);
            foreach (FooterPart footerPart in mainPart.FooterParts) UpdateMax(footerPart, ref max);
            UpdateMax(mainPart.FootnotesPart, ref max);
            UpdateMax(mainPart.EndnotesPart, ref max);
            UpdateMax(mainPart.WordprocessingCommentsPart, ref max);
            return max;
        }

        // A part nobody has loaded yet is parsed only when its markup can hold a drawing id at all, so a new document's
        // notes and comments stay unparsed.
        private static void UpdateMax(OpenXmlPart? part, ref uint max) {
            if (part == null) return;
            if (!part.IsRootElementLoaded && !MayHoldDrawingIds(part)) return;
            UpdateMax(part.RootElement, ref max);
        }

        private static readonly byte[] DocPrName = System.Text.Encoding.ASCII.GetBytes("docPr");
        private static readonly byte[] CNvPrName = System.Text.Encoding.ASCII.GetBytes("cNvPr");

        private static bool MayHoldDrawingIds(OpenXmlPart part) {
            byte[] bytes;
            using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read)) {
                using var buffer = new MemoryStream();
                stream.CopyTo(buffer);
                bytes = buffer.ToArray();
            }
            // Element names are searched as UTF-8; a part in another encoding is parsed.
            if (bytes.Length >= 2 && (bytes[0] == 0xFF || bytes[0] == 0xFE || bytes[0] == 0x00 || bytes[1] == 0x00)) return true;
            return IndexOf(bytes, DocPrName) >= 0 || IndexOf(bytes, CNvPrName) >= 0;
        }

        private static int IndexOf(byte[] data, byte[] value) {
#if NET
            return data.AsSpan().IndexOf(value);
#else
            for (int start = 0; start <= data.Length - value.Length; start++) {
                int offset = 0;
                while (offset < value.Length && data[start + offset] == value[offset]) offset++;
                if (offset == value.Length) return start;
            }
            return -1;
#endif
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
            internal bool Seeded { get; set; }
            internal bool Shared { get; set; }
        }
    }
}
