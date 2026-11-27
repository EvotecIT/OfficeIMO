using System.Collections.ObjectModel;

namespace OfficeIMO.Word.Pdf;


public static partial class WordPdfConverterExtensions {
    /// <summary>
    /// Boxes paint in the story's deferred layers, ordered by z-order across the whole header or footer,
    /// so only boxes that must not overlap need Word's repositioning.
    /// </summary>
    private readonly record struct NativeRunningTextBoxStoryPeers(int VisibleBoxCount, bool HasDisallowedOverlap) {
        public bool RequiresApproximation => VisibleBoxCount > 1 && HasDisallowedOverlap;
    }

    /// <summary>Captures each underlying header/footer part once while operation visibility scopes are active.</summary>
    private static IReadOnlyDictionary<OfficeXmlNode, NativeRunningTextBoxStoryPeers> GetNativeRunningTextBoxStoryPeers(
        WordDocument document, WordPdfExportOptions? options) {
        var result = new Dictionary<OfficeXmlNode, NativeRunningTextBoxStoryPeers>();
        foreach (OfficeXmlNode root in document.HeaderFooterRootNodes()) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            if (result.ContainsKey(root)) continue;
            int count = 0;
            bool disallowed = false;
            foreach (OfficeXmlNode source in root.Descendants(WordXml.W, "p")) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                // Nested textbox content belongs to its outer box, including
                // when that outer hosting run is hidden. It is not a story peer.
                if (!source.Ancestor(WordXml.W, "txbxContent").IsNone) continue;
                WordParagraph paragraph = WordConverterNodes.ParagraphView(document, source);
                foreach (WordTextBox box in GetNativeVisibleTextBoxes(paragraph, GetNativeRuns(paragraph))) {
                    options?.CancellationToken.ThrowIfCancellationRequested();
                    count = Math.Min(2, count + 1);
                    disallowed |= WordXml.ParseBoolean(box.AnchorNode().Attr("allowOverlap")) != true;
                }
            }
            result.Add(root, new NativeRunningTextBoxStoryPeers(count, disallowed));
        }
        return new ReadOnlyDictionary<OfficeXmlNode, NativeRunningTextBoxStoryPeers>(result);
    }

    private static bool HasNativeRunningStoryPeerConstraint(WordHeaderFooter? story, NativeDocumentDefaults defaults) {
        OfficeXmlNode root = story?.Node() ?? default;
        if (root.IsNone) return false;
        return defaults.RunningTextBoxStoryPeers == null ||
            !defaults.RunningTextBoxStoryPeers.TryGetValue(root, out NativeRunningTextBoxStoryPeers peers) ||
            peers.RequiresApproximation;
    }
}
