using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

internal static partial class OpenDocumentReaderAdapter {
    private static IEnumerable<ReaderChunk> ReadDrawing(OdgDocument document, string sourceName, ProjectionBudget budget, CancellationToken cancellationToken) {
        IReadOnlyList<OdgPage> pages = document.Pages;
        for (int index = 0; index < pages.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            OdgPage page = pages[index];
            string name = budget.Read(page.Name);
            var paragraphs = new List<string>();
            var pending = new Stack<IEnumerator<OdgShape>>();
            pending.Push(page.Shapes.GetEnumerator());
            try {
                while (pending.Count > 0) {
                    cancellationToken.ThrowIfCancellationRequested();
                    IEnumerator<OdgShape> current = pending.Peek();
                    if (!current.MoveNext()) { pending.Pop().Dispose(); continue; }
                    OdgShape shape = current.Current;
                    if (shape.IsGroup) pending.Push(shape.Children.GetEnumerator());
                    else if (!string.IsNullOrWhiteSpace(shape.Text)) paragraphs.Add(budget.Read(shape.Text));
                }
            } finally { foreach (var iterator in pending) iterator.Dispose(); }
            string text = string.Join(Environment.NewLine, paragraphs);
            yield return new ReaderChunk {
                Id = BuildId(sourceName, "page", index), Kind = ReaderInputKind.OpenDocument,
                Text = text, Markdown = "## Page " + (index + 1).ToString(CultureInfo.InvariantCulture) + ": " + name + Environment.NewLine + Environment.NewLine + text,
                Location = new ReaderLocation { Path = sourceName, Page = index + 1, BlockIndex = index, SourceBlockIndex = index, SourceBlockKind = "drawing-page", HeadingPath = name },
                Warnings = new[] { "Drawing extraction includes shape text in paint order, including hidden layers; master text, image OCR and embedded objects are not extracted." }
            };
        }
    }
}
