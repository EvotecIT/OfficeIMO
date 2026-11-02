using OfficeIMO.IWork;
using System.Threading;
using OpenXmlTable = DocumentFormat.OpenXml.Wordprocessing.Table;

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

    private static bool AddInlineObject(WordDocument document, WordParagraph paragraph,
        IWorkPagesDrawable drawable, IWorkNativeListCatalog nativeLists, double contentWidth, double contentHeight,
        List<WordCellComment> cellComments, HashSet<ulong> standaloneTables, IWorkPageLayout? pageLayout,
        CancellationToken cancellationToken) {
        if (drawable.Table is { } table) {
            WordTable? target = AddTable(document, table, nativeLists, paragraph, null, cellComments, cancellationToken);
            if (target != null && standaloneTables.Contains(DrawableIdentifier(drawable))) {
                if (CanApplyInlineTableGeometry(table, pageLayout, standaloneTables, cancellationToken))
                    ApplyInlineTableGrid(target, table, cancellationToken);
                // Word joins adjacent same-style tables without a paragraph boundary.
                // Keep this host when removing it would join distinct source tables.
                if (paragraph._paragraph.PreviousSibling() is not OpenXmlTable
                    && target._table.NextSibling() is not OpenXmlTable) {
                    paragraph.Remove();
                    return true;
                }
            }
        } else if (drawable.Image is { } source) {
            using var image = new MemoryStream(source.GetBytes(), writable: false);
            double width = source.Geometry?.WidthPoints ?? source.PixelWidth.GetValueOrDefault(640) * 72d / 96d;
            double height = source.Geometry?.HeightPoints ?? source.PixelHeight.GetValueOrDefault(480) * 72d / 96d;
            if (source.Geometry == null) (width, height) = FitInside(width, height, contentWidth, contentHeight);
            paragraph.AddText(string.Empty).AddImage(image, source.FileName, OfficeIMO.OfficeLength.Pixels(Helpers.ConvertPointsToPixels(width)), OfficeIMO.OfficeLength.Pixels(Helpers.ConvertPointsToPixels(height)), options: new OfficeIMO.Word.WordImageOptions { Description = source.AccessibilityDescription ?? "Image imported from Pages" });
        }
        return false;
    }
}
