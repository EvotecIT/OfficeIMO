using System.Collections.Generic;
using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private readonly record struct NativeRunningTextBoxStoryPeers(int VisibleBoxCount, bool HasDisallowedOverlap) {
        public bool RequiresApproximation => VisibleBoxCount > 1 && HasDisallowedOverlap;
    }

    /// <summary>Captures each underlying header/footer part once while operation visibility scopes are active.</summary>
    private static IReadOnlyDictionary<OpenXmlElement, NativeRunningTextBoxStoryPeers> GetNativeRunningTextBoxStoryPeers(
        WordDocument document, WordToPdfOptions? options) {
        var result = new Dictionary<OpenXmlElement, NativeRunningTextBoxStoryPeers>();
        var main = document._wordprocessingDocument.MainDocumentPart;
        if (main != null)
            foreach (OpenXmlElement root in main.HeaderParts.Select(part => (OpenXmlElement?)part.Header)
                    .Concat(main.FooterParts.Select(part => (OpenXmlElement?)part.Footer)).OfType<OpenXmlElement>()) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                if (result.ContainsKey(root)) continue;
                int count = 0;
                bool disallowed = false;
                foreach (W.Paragraph source in root.Descendants<W.Paragraph>()) {
                    options?.CancellationToken.ThrowIfCancellationRequested();
                    // Nested textbox content belongs to its outer box, including
                    // when that outer hosting run is hidden. It is not a story peer.
                    if (source.Ancestors<W.TextBoxContent>().Any()) continue;
                    var paragraph = new WordParagraph(document, source);
                    foreach (WordTextBox box in GetNativeVisibleTextBoxes(paragraph, GetNativeRuns(paragraph))) {
                        options?.CancellationToken.ThrowIfCancellationRequested();
                        count = Math.Min(2, count + 1);
                        disallowed |= box.Anchor?.AllowOverlap?.Value != true;
                    }
                }
                result.Add(root, new NativeRunningTextBoxStoryPeers(count, disallowed));
            }
        return new ReadOnlyDictionary<OpenXmlElement, NativeRunningTextBoxStoryPeers>(result);
    }

    private static bool HasNativeRunningStoryPeerConstraint(WordHeaderFooter? story, NativeDocumentDefaults defaults) {
        OpenXmlElement? root = (OpenXmlElement?)story?._header ?? story?._footer;
        if (root == null) return false;
        return defaults.RunningTextBoxStoryPeers == null ||
            !defaults.RunningTextBoxStoryPeers.TryGetValue(root, out NativeRunningTextBoxStoryPeers peers) ||
            peers.RequiresApproximation;
    }
}
