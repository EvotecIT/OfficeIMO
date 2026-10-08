using System.Globalization;
using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static List<(EpubMediaOverlayCue Cue, XElement Parallel)> PrepareOverlayNodes(EpubMediaOverlay overlay,
        XElement body, Dictionary<string, XElement> targets, XElement rootSequence, string overlayPath,
        string contentPath, CancellationToken token) {
        if (overlay.Cues == null || overlay.Nodes == null || overlay.Cues.Count > 10000 || overlay.Nodes.Count > 10000 ||
            (overlay.Cues.Count == 0) == (overlay.Nodes.Count == 0))
            throw new ArgumentException("Supply either Cues or Nodes, with one to 10,000 entries.", nameof(overlay));
        IReadOnlyList<EpubMediaOverlayNode> roots = overlay.Nodes.Count != 0 ? overlay.Nodes : overlay.Cues;
        var cues = new List<(EpubMediaOverlayCue, XElement)>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        int count = 0, sequenceId = 0;
        void Append(IReadOnlyList<EpubMediaOverlayNode> nodes, XElement parentTarget, XElement parent, int depth) {
            foreach (EpubMediaOverlayNode node in nodes) {
                token.ThrowIfCancellationRequested();
                if (++count > 10000) throw new ArgumentException("An overlay supports at most 10,000 total nodes.", nameof(overlay));
                if (node == null || node.ElementId.Length == 0 || node.ElementId.Length > 1024 ||
                    !targets.TryGetValue(node.ElementId, out XElement? target) || !IsNarratableElement(target) || !seen.Add(node.ElementId) ||
                    !(target.Ancestors().Contains(parentTarget) || depth == 0 && target == body))
                    throw new InvalidDataException("Narration nodes require distinct supported content targets contained by their parent sequence.");
                string? semantic = OverlaySemantic(node.Semantic);
                XElement element;
                if (node is EpubMediaOverlayCue cue) {
                    element = new XElement(Smil + "par");
                    cues.Add((cue, element));
                } else if (node is EpubMediaOverlaySequence sequence) {
                    if (depth >= 32) throw new ArgumentException("An overlay supports at most 32 nested sequences.", nameof(overlay));
                    element = new XElement(Smil + "seq", new XAttribute("id", "seq" + (sequenceId++).ToString(CultureInfo.InvariantCulture)),
                        new XAttribute(Ops + "textref", RelativeHref(overlayPath, contentPath) + "#" + Uri.EscapeDataString(node.ElementId)));
                    Append(sequence.Children, target, element, depth + 1);
                } else throw new NotSupportedException("Unsupported narration node.");
                if (semantic != null) element.SetAttributeValue(Ops + "type", semantic);
                parent.Add(element);
            }
        }
        Append(roots, body, rootSequence, 0);
        return cues;
    }

    private static string? OverlaySemantic(EpubMediaOverlaySemantic? semantic) {
        if (semantic == null) return null;
        switch (semantic.Value) {
            case EpubMediaOverlaySemantic.Footnote: return "footnote";
            case EpubMediaOverlaySemantic.Endnote: return "endnote";
            case EpubMediaOverlaySemantic.PageBreak: return "pagebreak";
            case EpubMediaOverlaySemantic.Table: return "table";
            case EpubMediaOverlaySemantic.TableRow: return "table-row";
            case EpubMediaOverlaySemantic.TableCell: return "table-cell";
            case EpubMediaOverlaySemantic.List: return "list";
            case EpubMediaOverlaySemantic.ListItem: return "list-item";
            case EpubMediaOverlaySemantic.Figure: return "figure";
            case EpubMediaOverlaySemantic.Aside: return "aside";
            default: throw new ArgumentOutOfRangeException(nameof(semantic));
        }
    }

    private static void RequireOverlaySemantic(XElement element) {
        string? value = (string?)element.Attribute(Ops + "type");
        if (value != null && !Enum.GetValues(typeof(EpubMediaOverlaySemantic)).Cast<EpubMediaOverlaySemantic>().Any(item => OverlaySemantic(item) == value))
            throw new NotSupportedException("Replacement cannot preserve an unsupported narration semantic.");
    }

    private static string OverlayNodeKey(XElement element, string path) {
        string href = element.Name == Smil + "par" ? (string)element.Element(Smil + "text")!.Attribute("src")! :
            (string)element.Attribute(Ops + "textref")!;
        return element.Name.LocalName + "#" + EpubReference.Resolve(path, href).Fragment;
    }
}
