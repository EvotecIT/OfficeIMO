using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    internal static PdfMutationOperation GetAttributeOperation(IReadOnlyList<PdfPageAttributeEdit> edits) =>
        edits.Count > 0 && edits.All(edit => edit is not null &&
            edit.Operation is PdfPageAttributeEdit.Kind.SetRotation or PdfPageAttributeEdit.Kind.RotateBy)
            ? PdfMutationOperation.RotatePages : PdfMutationOperation.ModifyPageAttributes;

    internal static byte[] EditPageAttributes(byte[] pdf, IReadOnlyList<PdfPageAttributeEdit> edits,
        PdfLoadOptions? readOptions, CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(edits, nameof(edits));
        if (edits.Count == 0) throw new ArgumentException("Supply at least one page-attribute edit.", nameof(edits));
        PdfMutationPlan plan = PdfMutationPlanner.Require(pdf, GetAttributeOperation(edits),
            readOptions, cancellationToken: cancellationToken);
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        var overrides = new Dictionary<int, Dictionary<string, PdfObject>>();
        foreach (PdfPageAttributeEdit edit in edits) {
            cancellationToken.ThrowIfCancellationRequested();
            Guard.NotNull(edit, nameof(edits));
            int[] pages = edit.Selection?.ToPageNumbers(document.Pages.Count, nameof(edits)) ??
                Enumerable.Range(1, document.Pages.Count).ToArray();
            foreach (int number in pages.Distinct()) {
                cancellationToken.ThrowIfCancellationRequested();
                PdfReadPage page = document.Pages[number - 1];
                if (!overrides.TryGetValue(page.ObjectNumber, out Dictionary<string, PdfObject>? values)) {
                    values = new Dictionary<string, PdfObject>(StringComparer.Ordinal);
                    overrides.Add(page.ObjectNumber, values);
                }
                ApplyPageAttributeEdit(page, values, edit);
            }
        }
        return WritePageAttributes(pdf, document, overrides, plan, readOptions, cancellationToken);
    }

    private static void ApplyPageAttributeEdit(PdfReadPage page, Dictionary<string, PdfObject> values,
        PdfPageAttributeEdit edit) {
        int rotation = values.TryGetValue("Rotate", out PdfObject? value)
            ? (int)((PdfNumber)value).Value : NormalizeRotation(page.GetRotationDegrees());
        if (edit.Operation == PdfPageAttributeEdit.Kind.SetRotation || edit.Operation == PdfPageAttributeEdit.Kind.RotateBy) {
            int delta = NormalizeRotation(edit.Rotation);
            values["Rotate"] = new PdfNumber(edit.Operation == PdfPageAttributeEdit.Kind.RotateBy
                ? NormalizeRotation(rotation + delta) : delta);
            return;
        }
        if (edit.Operation == PdfPageAttributeEdit.Kind.SetBox) {
            ValidatePageBoxCoordinates(edit.Left, edit.Bottom, edit.Right, edit.Top);
            values[GetPageBoxName(edit.Box)] = CreatePageBoxArray(edit.Left, edit.Bottom, edit.Right, edit.Top);
            return;
        }
        PdfPageGeometry geometry = page.GetGeometry();
        PdfPageBox media = GetBatchBox(values, "MediaBox", geometry.MediaBox) ??
            throw new InvalidOperationException("The page has no readable MediaBox.");
        if (edit.Operation == PdfPageAttributeEdit.Kind.ResetCrop) {
            values["CropBox"] = CreatePageBoxArray(media.Left, media.Bottom, media.Right, media.Top);
            return;
        }
        PdfPageBox crop = GetBatchBox(values, "CropBox", geometry.CropBox) ?? media;
        var effective = new PdfPageBox("EffectiveBox", Math.Max(media.Left, crop.Left), Math.Max(media.Bottom, crop.Bottom),
            Math.Min(media.Right, crop.Right), Math.Min(media.Top, crop.Top));
        if (effective.Width <= 0 || effective.Height <= 0)
            throw new InvalidOperationException("The page boxes have no visible intersection.");
        double[] margins = { edit.Left, edit.Bottom, edit.Right, edit.Top };
        if (margins.Any(margin => !IsFinite(margin) || margin < 0))
            throw new ArgumentOutOfRangeException(nameof(edit), "Crop margins must be finite, non-negative physical points.");
        double unit = geometry.UserUnit ?? 1D;
        var size = PdfVisualCoordinateMapper.GetVisualSize(effective, rotation, unit);
        if (edit.Left + edit.Right >= size.Width || edit.Bottom + edit.Top >= size.Height)
            throw new ArgumentOutOfRangeException(nameof(edit), "Crop margins must leave a positive visible page area.");
        PdfVisualBounds bounds = PdfVisualCoordinateMapper.TransformVisualBoundsToUser(effective, rotation,
            edit.Left, edit.Top, size.Width - edit.Right, size.Height - edit.Bottom, unit);
        // The inverse mapper returns min/max user-space Y in Top/Bottom respectively.
        values["CropBox"] = CreatePageBoxArray(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom);
    }

    private static PdfPageBox? GetBatchBox(Dictionary<string, PdfObject> values, string name, PdfPageBox? original) {
        if (!values.TryGetValue(name, out PdfObject? value)) return original;
        var array = (PdfArray)value;
        return new PdfPageBox(name, ((PdfNumber)array.Items[0]).Value, ((PdfNumber)array.Items[1]).Value,
            ((PdfNumber)array.Items[2]).Value, ((PdfNumber)array.Items[3]).Value);
    }
}
