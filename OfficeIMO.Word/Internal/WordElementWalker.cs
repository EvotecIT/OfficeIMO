using DocumentFormat.OpenXml;

namespace OfficeIMO.Word;

/// <summary>Answers existence questions over an element subtree without allocating a descendant enumerator.</summary>
internal static class WordElementWalker {
    /// <summary>Whether <paramref name="root"/> or any of its descendants satisfies <paramref name="match"/>, in document order.</summary>
    internal static bool SubtreeContains(OpenXmlElement root, Func<OpenXmlElement, bool> match) =>
        match(root) || DescendantsContain(root, match);

    /// <summary>Whether <paramref name="run"/> has a comment reference among its direct children.</summary>
    internal static bool HasDirectCommentReference(DocumentFormat.OpenXml.Wordprocessing.Run run) {
        for (OpenXmlElement? child = run.FirstChild; child != null; child = child.NextSibling()) {
            if (child is DocumentFormat.OpenXml.Wordprocessing.CommentReference) return true;
        }
        return false;
    }

    /// <summary>Whether any descendant of <paramref name="root"/> satisfies <paramref name="match"/>, in document order.</summary>
    internal static bool DescendantsContain(OpenXmlElement root, Func<OpenXmlElement, bool> match) {
        OpenXmlElement? current = root.FirstChild;
        while (current != null) {
            if (match(current)) return true;
            OpenXmlElement? next = current.FirstChild;
            while (next == null) {
                next = current.NextSibling();
                if (next != null) break;
                current = current.Parent;
                if (current == null || ReferenceEquals(current, root)) return false;
            }
            current = next;
        }
        return false;
    }
}
