using System.Globalization;

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
        /// or <see cref="OfficeXmlNode.None"/> when the current view omits it.
        /// </summary>
        private static OfficeXmlNode GetNativeVisibleRevisionContainer(OfficeXmlNode element) {
            if (element.IsTrackedChangeContainer() && element.LocalName is "ins" or "moveTo") return element;
            if (element.IsTrackedChangeContainer() && element.LocalName is "del" or "moveFrom" &&
                CurrentNativeTrackedChangeView == WordPdfTrackedChangeView.Markup) return element;
            return default;
        }

        /// <summary>
        /// Applies the markup view's revision formatting to a run inside a tracked change.
        /// Returns the source run unchanged in the final view or outside revisions.
        /// </summary>
        private static OfficeXmlNode ApplyNativeTrackedChangeMarkup(OfficeXmlNode sourceRun) {
            if (CurrentNativeTrackedChangeView != WordPdfTrackedChangeView.Markup) return sourceRun;
            bool deleted = false;
            bool inserted = false;
            for (OfficeXmlNode ancestor = sourceRun.Parent; !ancestor.IsNone && !ancestor.IsW("p"); ancestor = ancestor.Parent) {
                if (ancestor.IsW("del") || ancestor.IsW("moveFrom")) deleted = true;
                else if (ancestor.IsW("ins") || ancestor.IsW("moveTo")) inserted = true;
            }
            if (!deleted && !inserted) return sourceRun;

            OfficeXmlNode markupRun = sourceRun.Clone();
            if (deleted) {
                foreach (OfficeXmlNode deletedText in markupRun.WElements("delText").ToList()) {
                    OfficeXmlNode text = markupRun.CreateElement(WordXml.W, "t");
                    text.SetInnerText(deletedText.InnerText);
                    text.SetAttribute(WordXml.Xml, "space", "preserve");
                    deletedText.InsertAfterSelf(text);
                    deletedText.Remove();
                }
            }
            OfficeXmlNode properties = markupRun.GetOrPrependChild(WordXml.W, "rPr");
            properties.RemoveChildren(WordXml.W, "color");
            AppendNativeMarkupProperty(properties, "color", NativeTrackedChangeMarkupColor);
            if (deleted) {
                properties.RemoveChildren(WordXml.W, "strike");
                properties.RemoveChildren(WordXml.W, "dstrike");
                AppendNativeMarkupProperty(properties, "strike", null);
            } else {
                properties.RemoveChildren(WordXml.W, "u");
                AppendNativeMarkupProperty(properties, "u", "single");
            }
            return markupRun;
        }

        private static void AppendNativeMarkupProperty(OfficeXmlNode properties, string localName, string? value) {
            OfficeXmlNode property = properties.CreateElement(WordXml.W, localName);
            if (value != null) property.SetAttribute(WordXml.W, "val", value);
            properties.Append(property);
        }

        private static void ReportNativeTrackedChanges(WordDocument document, WordPdfExportOptions options) {
            if (options.TrackedChanges != WordPdfTrackedChangeView.Final) return;
            if (document.DocumentNode().IsNone) return;
            int insertions = 0;
            int deletions = 0;
            foreach (OfficeXmlNode root in EnumerateNativeRevisionRoots(document)) {
                foreach (OfficeXmlNode element in root.DescendantElements()) {
                    if (element.IsTrackedChangeContainer() && element.LocalName is "ins" or "moveTo") insertions++;
                    else if (element.IsTrackedChangeContainer() && element.LocalName is "del" or "moveFrom") deletions++;
                }
            }
            if (insertions == 0 && deletions == 0) return;

            string counts = insertions.ToString(CultureInfo.InvariantCulture) + " insertion(s) and " +
                deletions.ToString(CultureInfo.InvariantCulture) + " deletion(s)";
            string message = "Tracked changes (" + counts + ") are shown as the final document: insertions are current text and deleted text is omitted. Set TrackedChanges to Markup to show the revisions.";
            AddNativeExportWarning(options, "tracked-revisions-projected", "tracked changes", message);
        }

        private static IEnumerable<OfficeXmlNode> EnumerateNativeRevisionRoots(WordDocument document) {
            if (document.BodyNode() is { Exists: true } body) yield return body;
            foreach (OfficeXmlNode root in document.HeaderFooterRootNodes()) yield return root;
            if (document.FootnotesNode() is { Exists: true } footnotes) yield return footnotes;
            if (document.EndnotesNode() is { Exists: true } endnotes) yield return endnotes;
        }
    }
}
