using OfficeIMO.IWork;
using System.Threading;

namespace OfficeIMO.Word.IWork;

public static partial class WordIWorkConverter {
    private static ulong DrawableIdentifier(IWorkPagesDrawable drawable) =>
        (drawable.Image?.SourceIdentity ?? drawable.Table?.SourceIdentity ?? drawable.TextBox?.SourceIdentity)!.RecordIdentifier;

    private static string? FindInlineObjectLimitation(IWorkPagesProjection projection) {
        var drawables = projection.Drawables.ToDictionary(DrawableIdentifier);
        var seen = new HashSet<ulong>();
        foreach (IWorkTextParagraph paragraph in projection.Body.Paragraphs) {
            foreach (IWorkTextRun run in paragraph.Runs.Where(run => run.InlineObject != null)) {
                ulong identifier = run.InlineObject!.Drawable.RecordIdentifier;
                if (!drawables.TryGetValue(identifier, out IWorkPagesDrawable? drawable)
                    || drawable.Kind is not (IWorkPagesDrawableKind.Image or IWorkPagesDrawableKind.Table))
                    return "A Pages inline attachment has no supported projected image or table.";
                if (!seen.Add(identifier)) return "A Pages drawable has more than one inline attachment position.";
                if (run.Hyperlink != null) return "A Pages inline attachment carries a hyperlink that the DOCX owner cannot preserve.";
                if (drawable.Kind == IWorkPagesDrawableKind.Table
                    && (paragraph.Text.Length != 0 || paragraph.Runs.Count(other => other.InlineObject != null) != 1))
                    return "A Pages inline table shares a paragraph with other content; DOCX requires a separate table block.";
            }
        }
        return null;
    }

    private static void AddInlineObject(WordDocument document, WordParagraph paragraph,
        IWorkPagesDrawable drawable, IWorkNativeListCatalog nativeLists, double contentWidth, double contentHeight,
        CancellationToken cancellationToken) {
        if (drawable.Table is { } table) {
            AddTable(document, table, nativeLists, paragraph, null, cancellationToken);
        } else if (drawable.Image is { } source) {
            using var image = new MemoryStream(source.GetBytes(), writable: false);
            double width = source.Geometry?.WidthPoints ?? source.PixelWidth.GetValueOrDefault(640) * 72d / 96d;
            double height = source.Geometry?.HeightPoints ?? source.PixelHeight.GetValueOrDefault(480) * 72d / 96d;
            if (source.Geometry == null) (width, height) = FitInside(width, height, contentWidth, contentHeight);
            paragraph.AddText(string.Empty).InsertImage(image, source.FileName, width, height,
                description: source.AccessibilityDescription ?? "Image imported from Pages");
        }
    }
}
