namespace OfficeIMO.OneNote;

/// <summary>Centralizes recursive traversal of the current semantic page content tree.</summary>
internal static class OneNoteElementTraversal {
    internal static IEnumerable<OneNoteElement> Enumerate(OneNotePage page) {
        ArgumentNullException.ThrowIfNull(page);
        foreach (OneNoteElement element in page.DirectContent) {
            foreach (OneNoteElement nested in Enumerate(element)) yield return nested;
        }
        foreach (OneNoteOutline outline in page.Outlines) {
            foreach (OneNoteElement nested in Enumerate(outline)) yield return nested;
        }
    }

    internal static IEnumerable<OneNoteElement> Enumerate(OneNoteElement element) {
        ArgumentNullException.ThrowIfNull(element);
        yield return element;
        if (element is OneNoteOutline outline) {
            foreach (OneNoteElement child in outline.Children) {
                foreach (OneNoteElement nested in Enumerate(child)) yield return nested;
            }
        } else if (element is OneNoteParagraph paragraph) {
            foreach (OneNoteElement child in paragraph.Children) {
                foreach (OneNoteElement nested in Enumerate(child)) yield return nested;
            }
        } else if (element is OneNoteTable table) {
            foreach (OneNoteTableRow row in table.Rows)
            foreach (OneNoteTableCell cell in row.Cells)
            foreach (OneNoteElement child in cell.Content)
            foreach (OneNoteElement nested in Enumerate(child)) yield return nested;
        }
    }
}
