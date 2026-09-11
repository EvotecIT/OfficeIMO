using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    internal Task TransformSelectedObjectAsync(PdfObjectTransformGesture gesture, long revision, PdfImageEditLayer imageLayer,
        CancellationToken token, IProgress<PdfWorkspaceProgress>? progress) {
        var selection = gesture.Selection;
        PdfDocument Current(byte[] bytes) {
            if (Revision != revision) throw new InvalidOperationException("The document changed during the object gesture. Select the object again.");
            return LoadDocument(bytes);
        }
        if (selection.Kind == PdfEditorSelectionKind.Annotation && selection.ObjectNumber is int number) {
            return MutateAnnotationBytesAsync(PdfWorkspaceOperationKind.Annotation, "Changed annotation geometry",
                [selection.PageNumber], bytes => {
                    PdfDocument document = Current(bytes);
                    PdfLogicalPage page = document.Read(new PdfReadOptions { Profile = PdfReadProfile.Fast }).Pages[selection.PageNumber - 1];
                    PdfEditorVisualBounds target = gesture.Target;
                    PdfPageRectangle bounds = page.MapVisualRectangleToUserSpace(target.Left, target.Top, target.Right, target.Bottom);
                    return document.Annotations.Resize(number, bounds);
                }, token, progress);
        }
        PdfImagePlacement placement = RequireImagePlacement(selection);
        return MutateBytesAsync(PdfWorkspaceOperationKind.ImageEdit, "Moved or resized image", [selection.PageNumber], bytes => {
            PdfDocument document = Current(bytes);
            PdfLogicalPage page = document.Read(new PdfReadOptions { Profile = PdfReadProfile.Fast }).Pages[selection.PageNumber - 1];
            PdfEditorVisualBounds original = selection.Bounds, target = gesture.Target;
            PdfPagePoint originalCenter = page.MapVisualPointToUserSpace((original.Left + original.Right) / 2, (original.Top + original.Bottom) / 2);
            PdfPagePoint targetCenter = page.MapVisualPointToUserSpace((target.Left + target.Right) / 2, (target.Top + target.Bottom) / 2);
            double scale = target.Width / original.Width;
            if (Math.Abs(target.Height / original.Height - scale) > 0.001) throw new InvalidOperationException("Image resizing must preserve its proportions.");
            return document.Images.Transform(placement, targetCenter.X - originalCenter.X, targetCenter.Y - originalCenter.Y,
                scale, new PdfImageEditOptions { Layer = imageLayer }).Document.ToBytes();
        }, token, progress);
    }
}
