using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    private static void NormalizePageAnnotations(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page,
        PageResizeTransform transform, ref int nextObject, CancellationToken cancellationToken) {
        if (!page.Items.TryGetValue("Annots", out PdfObject? annotationsValue)) return;
        if (PdfObjectLookup.Resolve(objects, annotationsValue) is not PdfArray annotations)
            throw new InvalidOperationException("The page annotations cannot be resolved for rotation normalization.");
        foreach (PdfObject value in annotations.Items) {
            cancellationToken.ThrowIfCancellationRequested();
            if (PdfObjectLookup.Resolve(objects, value) is not PdfDictionary annotation ||
                !annotation.Items.TryGetValue("Rect", out PdfObject? rectangle))
                throw new InvalidOperationException("An annotation has no readable rectangle for rotation normalization.");
            double[] before = ReadNormalizationBox(objects, rectangle, allowEmpty: true);
            Matrix2D matrix = ToMatrix(transform);
            int flags = annotation.Get<PdfNumber>("F") is PdfNumber flagNumber ? (int)flagNumber.Value : 0;
            bool noRotate = (flags & 16) != 0;
            if (noRotate) {
                var anchor = matrix.Transform(before[0], before[3]);
                matrix = Matrix2D.Translation(anchor.X - before[0], anchor.Y - before[3]);
            }
            PdfArray after = TransformNormalizationBox(before[0], before[1], before[2], before[3], matrix);
            double left = ((PdfNumber)after.Items[0]).Value, bottom = ((PdfNumber)after.Items[1]).Value;
            double width = ((PdfNumber)after.Items[2]).Value - left, height = ((PdfNumber)after.Items[3]).Value - bottom;
            string? subtype = annotation.Get<PdfName>("Subtype")?.Name;
            bool empty = before[0] == before[2] || before[1] == before[3];
            if (empty) {
                // Zero-area widgets are valid, including invisible unsigned signature fields.
                // Their appearances do not paint and must not acquire a zero-sized wrapper.
            } else if (annotation.Items.TryGetValue("AP", out PdfObject? appearanceValue)) {
                if (PdfObjectLookup.Resolve(objects, appearanceValue) is not PdfDictionary appearance)
                    throw new InvalidOperationException("The annotation appearance dictionary cannot be resolved.");
                var rewritten = new PdfDictionary();
                foreach (var item in appearance.Items) {
                    rewritten.Items[item.Key] = NormalizeAppearanceStates(objects, item.Value, before, matrix,
                        left, bottom, width, height, ref nextObject, cancellationToken);
                }
                annotation.Items["AP"] = rewritten;
            } else if (subtype != "Link" && subtype != "Popup" && (flags & 3) == 0) {
                throw new NotSupportedException("Rotation normalization requires an appearance for visible " + subtype + " annotations.");
            }
            annotation.Items["Rect"] = after;
            var geometryTransform = new PageResizeTransform(matrix.A, matrix.B, matrix.C, matrix.D, matrix.E, matrix.F,
                transform.PageObjectNumber, transform.SourceLeft, transform.SourceBottom, transform.SourceWidth, transform.SourceHeight,
                transform.RotationDegrees, left, bottom, width, height);
            TransformAnnotationCoordinateArrays(objects, annotation, geometryTransform);
            TransformPairedNumberArray(objects, annotation, "CL", geometryTransform);
            if (!noRotate) {
                if (annotation.Items.TryGetValue("RD", out PdfObject? insetValue) &&
                    PdfObjectLookup.Resolve(objects, insetValue) is PdfArray insets && insets.Items.Count == 4) {
                    int shift = transform.RotationDegrees / 90;
                    var rotatedInsets = new PdfArray();
                    for (int i = 0; i < 4; i++) rotatedInsets.Items.Add(ClonePdfObject(insets.Items[(i + shift) % 4]));
                    annotation.Items["RD"] = rotatedInsets;
                }
                if (annotation.Items.TryGetValue("MK", out PdfObject? characteristicsValue) &&
                    PdfObjectLookup.Resolve(objects, characteristicsValue) is PdfDictionary characteristics) {
                    var cloned = (PdfDictionary)ClonePdfObject(characteristics);
                    int angle = (int)(cloned.Get<PdfNumber>("R")?.Value ?? 0);
                    cloned.Items["R"] = new PdfNumber(NormalizeRotation(angle - transform.RotationDegrees));
                    annotation.Items["MK"] = cloned;
                }
                if (subtype == "FreeText") {
                    int angle = (int)(annotation.Get<PdfNumber>("Rotate")?.Value ?? 0);
                    annotation.Items["Rotate"] = new PdfNumber(NormalizeRotation(angle - transform.RotationDegrees));
                }
            }
        }
    }

    private static PdfObject NormalizeAppearanceStates(Dictionary<int, PdfIndirectObject> objects, PdfObject value,
        double[] rectangle, Matrix2D pageMatrix, double left, double bottom, double width, double height,
        ref int nextObject, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfObject? resolved = PdfObjectLookup.Resolve(objects, value);
        if (resolved is PdfDictionary states) {
            var result = new PdfDictionary();
            foreach (var state in states.Items) {
                if (PdfObjectLookup.Resolve(objects, state.Value) is not PdfStream)
                    throw new InvalidOperationException("An annotation appearance state is not a stream.");
                result.Items[state.Key] = NormalizeAppearanceStates(objects, state.Value, rectangle, pageMatrix,
                    left, bottom, width, height, ref nextObject, cancellationToken);
            }
            return result;
        }
        if (resolved is not PdfStream stream)
            throw new InvalidOperationException("The annotation appearance cannot be resolved to a stream.");
        if (!PdfAppearancePlacement.TryRead(stream.Dictionary, item => PdfObjectLookup.Resolve(objects, item),
            rectangle[0], rectangle[1], rectangle[2] - rectangle[0], rectangle[3] - rectangle[1], out Matrix2D placement, out _))
            return value; // An empty appearance paints nothing, including an empty state on a visible widget.
        Matrix2D wrapperMatrix = Matrix2D.Multiply(Matrix2D.Translation(-left, -bottom), Matrix2D.Multiply(pageMatrix, placement));
        PdfReference original = value as PdfReference ?? new PdfReference(AddNormalizationObject(objects, ref nextObject, stream), 0);
        var xObjects = new PdfDictionary(); xObjects.Items["Original"] = original;
        var resources = new PdfDictionary(); resources.Items["XObject"] = xObjects;
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject"); dictionary.Items["Subtype"] = new PdfName("Form");
        dictionary.Items["BBox"] = CreatePageBoxArray(0, 0, width, height);
        dictionary.Items["Resources"] = resources;
        int wrapper = AddNormalizationObject(objects, ref nextObject, new PdfStream(dictionary,
            Encoding.ASCII.GetBytes("q\n" + NormalizationMatrixCommand(wrapperMatrix) + "/Original Do\nQ\n")));
        return new PdfReference(wrapper, 0);
    }
}
