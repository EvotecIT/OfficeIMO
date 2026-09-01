using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    internal Task ReplaceSelectedTextAsync(
        PdfEditorSelection selection,
        string replacement,
        PdfTextEditOptions? options,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        RequireSelection(selection, PdfEditorSelectionKind.Text);
        return MutateBytesAsync(
            PdfWorkspaceOperationKind.TextEdit,
            "Replaced selected text on page " + selection.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
            new[] { selection.PageNumber },
            bytes => {
                PdfDocument document = PdfDocument.Load(bytes);
                PdfTextMatch match = ResolveTextMatch(document, selection);
                return document.Text.Replace(match, replacement ?? string.Empty, options).Document.ToBytes();
            },
            cancellationToken,
            progress);
    }

    internal Task MoveSelectedTextAsync(
        PdfEditorSelection selection,
        double deltaX,
        double deltaY,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        RequireSelection(selection, PdfEditorSelectionKind.Text);
        return MutateBytesAsync(
            PdfWorkspaceOperationKind.TextEdit,
            "Moved selected text on page " + selection.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
            new[] { selection.PageNumber },
            bytes => {
                PdfDocument document = PdfDocument.Load(bytes);
                PdfTextMatch match = ResolveTextMatch(document, selection);
                return document.Text.Move(match, deltaX, deltaY).Document.ToBytes();
            },
            cancellationToken,
            progress);
    }

    internal Task ReplaceAllTextAsync(
        string search,
        string replacement,
        bool matchCase,
        bool wholeWords,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        if (string.IsNullOrEmpty(search)) throw new ArgumentException("Find text is required.", nameof(search));
        return MutateBytesAsync(
            PdfWorkspaceOperationKind.TextEdit,
            "Replaced all matching document text",
            Array.Empty<int>(),
            bytes => PdfDocument.Load(bytes).Text.ReplaceAll(
                search,
                replacement ?? string.Empty,
                new PdfTextSearchOptions { MatchCase = matchCase, WholeWords = wholeWords }).Document.ToBytes(),
            cancellationToken,
            progress);
    }

    internal Task MoveSelectedImageAsync(
        PdfEditorSelection selection,
        double deltaX,
        double deltaY,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        PdfImagePlacement placement = RequireImagePlacement(selection);
        return MutateBytesAsync(
            PdfWorkspaceOperationKind.ImageEdit,
            "Moved image on page " + selection.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
            new[] { selection.PageNumber },
            bytes => PdfDocument.Load(bytes).Images.Move(placement, deltaX, deltaY).Document.ToBytes(),
            cancellationToken,
            progress);
    }

    internal Task ReplaceSelectedImageAsync(
        PdfEditorSelection selection,
        byte[] imageBytes,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        PdfImagePlacement placement = RequireImagePlacement(selection);
        ArgumentNullException.ThrowIfNull(imageBytes);
        return MutateBytesAsync(
            PdfWorkspaceOperationKind.ImageEdit,
            "Replaced image on page " + selection.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
            new[] { selection.PageNumber },
            bytes => PdfDocument.Load(bytes).Images.Replace(placement, imageBytes).Document.ToBytes(),
            cancellationToken,
            progress);
    }

    internal Task RemoveSelectedImageAsync(
        PdfEditorSelection selection,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        PdfImagePlacement placement = RequireImagePlacement(selection);
        return MutateBytesAsync(
            PdfWorkspaceOperationKind.ImageEdit,
            "Removed image from page " + selection.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
            new[] { selection.PageNumber },
            bytes => PdfDocument.Load(bytes).Images.Remove(placement).Document.ToBytes(),
            cancellationToken,
            progress);
    }

    internal Task MoveAnnotationAsync(
        int objectNumber,
        int pageNumber,
        double deltaX,
        double deltaY,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(
            PdfWorkspaceOperationKind.Annotation,
            "Moved annotation on page " + pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
            new[] { pageNumber },
            bytes => PdfDocument.Load(bytes).Annotations.Move(objectNumber, deltaX, deltaY).Bytes,
            cancellationToken,
            progress);

    internal Task ResizeAnnotationAsync(
        int objectNumber,
        int pageNumber,
        PdfPageRectangle rectangle,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(
            PdfWorkspaceOperationKind.Annotation,
            "Resized annotation on page " + pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
            new[] { pageNumber },
            bytes => PdfDocument.Load(bytes).Annotations.Resize(objectNumber, rectangle).Bytes,
            cancellationToken,
            progress);

    private static PdfTextMatch ResolveTextMatch(PdfDocument document, PdfEditorSelection selection) {
        string text = selection.Text ?? throw new InvalidOperationException("The selected text is no longer available.");
        if (text.Length == 0) throw new InvalidOperationException("The selected text is empty.");
        PdfPageRegion region = CreatePageRegion(document, selection);
        PdfTextMatch[] candidates = document.Text.Find(
                text,
                new PdfTextSearchOptions { MatchCase = true, PageNumbers = new[] { selection.PageNumber } })
            .Where(match => Intersects(match, region))
            .OrderBy(match => DistanceSquared(match, region))
            .ToArray();
        return candidates.FirstOrDefault()
            ?? throw new InvalidOperationException("The selected text no longer matches an editable source occurrence.");
    }

    private static PdfPageRegion CreatePageRegion(PdfDocument document, PdfEditorSelection selection) {
        PdfLogicalPage page = document
            .Read(new PdfReadOptions { Profile = PdfReadProfile.Fast })
            .Pages.Single(candidate => candidate.PageNumber == selection.PageNumber);
        PdfEditorVisualBounds visual = selection.Bounds;
        PdfPageRectangle mapped = page.MapVisualRectangleToUserSpace(visual.Left, visual.Top, visual.Right, visual.Bottom);
        PdfPageBox? boundary = page.CropBox ?? page.MediaBox;
        return new PdfPageRegion(
            selection.PageNumber,
            mapped.Left - (boundary?.Left ?? 0D),
            mapped.Bottom - (boundary?.Bottom ?? 0D),
            mapped.Width,
            mapped.Height);
    }

    private static bool Intersects(PdfTextMatch match, PdfPageRegion region) =>
        match.X < region.Right && match.X + match.Width > region.X &&
        match.Y < region.Top && match.Y + match.Height > region.Y;

    private static double DistanceSquared(PdfTextMatch match, PdfPageRegion region) {
        double deltaX = match.X + (match.Width / 2D) - (region.X + (region.Width / 2D));
        double deltaY = match.Y + (match.Height / 2D) - (region.Y + (region.Height / 2D));
        return (deltaX * deltaX) + (deltaY * deltaY);
    }

    private static PdfImagePlacement RequireImagePlacement(PdfEditorSelection selection) {
        RequireSelection(selection, PdfEditorSelectionKind.Image);
        return selection.ImagePlacement
            ?? throw new InvalidOperationException("The selected image no longer has an exact placement identity.");
    }

    private static void RequireSelection(PdfEditorSelection selection, PdfEditorSelectionKind kind) {
        ArgumentNullException.ThrowIfNull(selection);
        if (selection.Kind != kind) throw new ArgumentException("The selected object kind does not match this operation.", nameof(selection));
    }
}
