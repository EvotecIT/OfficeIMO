using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        // Word prints every reviewer's revisions in an author colour; one
        // accessible red keeps the static markup readable on white paper.
        private const string NativeTrackedChangeMarkupColor = "C00000";

        private static readonly AsyncLocal<WordPdfTrackedChangeView?> NativeTrackedChangeViewScope = new();

        private static WordPdfTrackedChangeView CurrentNativeTrackedChangeView =>
            NativeTrackedChangeViewScope.Value ?? WordPdfTrackedChangeView.Markup;

        private static IDisposable UseNativeTrackedChangeView(WordPdfTrackedChangeView view) {
            WordPdfTrackedChangeView? previous = NativeTrackedChangeViewScope.Value;
            NativeTrackedChangeViewScope.Value = view;
            return new NativeTrackedChangeViewRestore(previous);
        }

        private sealed class NativeTrackedChangeViewRestore : IDisposable {
            private readonly WordPdfTrackedChangeView? _previous;
            private bool _disposed;

            internal NativeTrackedChangeViewRestore(WordPdfTrackedChangeView? previous) => _previous = previous;

            public void Dispose() {
                if (_disposed) return;
                _disposed = true;
                NativeTrackedChangeViewScope.Value = _previous;
            }
        }

        /// <summary>
        /// Returns the revision container a visible-run traversal should enter,
        /// or null when the current view omits it.
        /// </summary>
        private static OpenXmlCompositeElement? GetNativeVisibleRevisionContainer(OpenXmlElement element) =>
            element switch {
                W.InsertedRun inserted => inserted,
                W.MoveToRun movedTo => movedTo,
                W.DeletedRun deleted when CurrentNativeTrackedChangeView == WordPdfTrackedChangeView.Markup => deleted,
                W.MoveFromRun movedFrom when CurrentNativeTrackedChangeView == WordPdfTrackedChangeView.Markup => movedFrom,
                _ => null
            };

        /// <summary>
        /// Applies the markup view's revision formatting to a run inside a tracked change.
        /// Returns the source run unchanged in the final view or outside revisions.
        /// </summary>
        private static W.Run ApplyNativeTrackedChangeMarkup(W.Run sourceRun) {
            if (CurrentNativeTrackedChangeView != WordPdfTrackedChangeView.Markup) return sourceRun;
            bool deleted = false;
            bool inserted = false;
            for (OpenXmlElement? ancestor = sourceRun.Parent; ancestor != null && ancestor is not W.Paragraph; ancestor = ancestor.Parent) {
                if (ancestor is W.DeletedRun or W.MoveFromRun) deleted = true;
                else if (ancestor is W.InsertedRun or W.MoveToRun) inserted = true;
            }
            if (!deleted && !inserted) return sourceRun;

            var markupRun = (W.Run)sourceRun.CloneNode(true);
            if (deleted) {
                foreach (W.DeletedText deletedText in markupRun.Elements<W.DeletedText>().ToList()) {
                    var text = new W.Text(deletedText.Text) { Space = SpaceProcessingModeValues.Preserve };
                    deletedText.InsertAfterSelf(text);
                    deletedText.Remove();
                }
            }
            W.RunProperties properties = markupRun.RunProperties ?? markupRun.PrependChild(new W.RunProperties());
            properties.RemoveAllChildren<W.Color>();
            properties.Append(new W.Color { Val = NativeTrackedChangeMarkupColor });
            if (deleted) {
                properties.RemoveAllChildren<W.Strike>();
                properties.RemoveAllChildren<W.DoubleStrike>();
                properties.Append(new W.Strike());
            } else {
                properties.RemoveAllChildren<W.Underline>();
                properties.Append(new W.Underline { Val = W.UnderlineValues.Single });
            }
            return markupRun;
        }

        private static void ReportNativeTrackedChanges(WordDocument document, WordPdfExportOptions options) {
            if (options.TrackedChanges != WordPdfTrackedChangeView.Final) return;
            MainDocumentPart? mainPart = document._wordprocessingDocument?.MainDocumentPart;
            if (mainPart == null) return;
            int insertions = 0;
            int deletions = 0;
            foreach (OpenXmlElement root in EnumerateNativeRevisionRoots(mainPart)) {
                foreach (OpenXmlElement element in root.Descendants()) {
                    if (element is W.InsertedRun or W.MoveToRun) insertions++;
                    else if (element is W.DeletedRun or W.MoveFromRun) deletions++;
                }
            }
            if (insertions == 0 && deletions == 0) return;

            string counts = insertions.ToString(CultureInfo.InvariantCulture) + " insertion(s) and " +
                deletions.ToString(CultureInfo.InvariantCulture) + " deletion(s)";
            string message = "Tracked changes (" + counts + ") are shown as the final document: insertions are current text and deleted text is omitted. Set TrackedChanges to Markup to show the revisions.";
            AddNativeExportWarning(options, "tracked-revisions-projected", "tracked changes", message);
        }

        private static IEnumerable<OpenXmlElement> EnumerateNativeRevisionRoots(MainDocumentPart mainPart) {
            if (mainPart.Document?.Body is { } body) yield return body;
            foreach (HeaderPart header in mainPart.HeaderParts)
                if (header.Header != null) yield return header.Header;
            foreach (FooterPart footer in mainPart.FooterParts)
                if (footer.Footer != null) yield return footer.Footer;
            if (mainPart.FootnotesPart?.Footnotes is { } footnotes) yield return footnotes;
            if (mainPart.EndnotesPart?.Endnotes is { } endnotes) yield return endnotes;
        }
    }
}
