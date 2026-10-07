using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>
    /// Applies up to 10,000 non-overlapping element replacements/deletions atomically. Expected XML
    /// detects stale drafts. XHTML targets must be inside body; SVG targets must be below its root.
    /// Identifiers and publication references are validated against the complete proposed batch.
    /// </summary>
    public void ApplyContentEdits(IEnumerable<EpubContentEdit> edits, CancellationToken cancellationToken = default) {
        if (edits == null) throw new ArgumentNullException(nameof(edits));
        cancellationToken.ThrowIfCancellationRequested();
        var documents = new Dictionary<string, XDocument>(StringComparer.Ordinal);
        var targets = new List<(XElement Target, XElement? Replacement)>();
        var selected = new HashSet<XElement>();
        foreach (EpubContentEdit edit in edits) {
            cancellationToken.ThrowIfCancellationRequested();
            if (targets.Count == 10_000) throw new InvalidDataException("Scoped content edits exceed the operation-count bound.");
            if (edit == null) throw new ArgumentException("Content edits cannot contain null entries.", nameof(edits));
            if (!documents.TryGetValue(edit.ManifestId, out XDocument? document)) {
                EpubManifestItem item = RequireManifestItem(edit.ManifestId);
                document = GetContentXml(edit.ManifestId);
                ValidateContent(document, item.MediaType);
                documents.Add(edit.ManifestId, document);
            }
            XElement target = RequireContentElement(document, edit.ElementId);
            bool inBody = document.Root!.Name == Html + "html" && target.Ancestors(Html + "body").Any(body => body.Parent == document.Root);
            bool inSvg = document.Root.Name == XName.Get("svg", "http://www.w3.org/2000/svg") && target != document.Root;
            if (!inBody && !inSvg) throw new NotSupportedException("Scoped edits select body content or SVG descendants, not document scaffolding.");
            if (!XNode.DeepEquals(target, edit.Expected)) throw new InvalidOperationException("The selected element has changed since the edit was prepared: " + edit.ElementId);
            if (!selected.Add(target))
                throw new ArgumentException("Scoped edits cannot select the same element or overlapping ancestors and descendants.", nameof(edits));
            targets.Add((target, edit.Replacement));
        }
        foreach (var edit in targets)
            if (edit.Target.Ancestors().Any(selected.Contains)) throw new ArgumentException("Scoped edits cannot overlap ancestors and descendants.", nameof(edits));
        foreach (var edit in targets) {
            cancellationToken.ThrowIfCancellationRequested();
            if (edit.Replacement == null) edit.Target.Remove(); else edit.Target.ReplaceWith(edit.Replacement);
        }
        SetContentXml(documents, cancellationToken);
    }
}
