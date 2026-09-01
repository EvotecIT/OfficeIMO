namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationEditor {
    internal static PdfAnnotationEditResult MoveAnnotation(
        byte[] pdf,
        int objectNumber,
        double deltaX,
        double deltaY,
        PdfLoadOptions? readOptions) {
        ValidateTransformFinite(deltaX, nameof(deltaX));
        ValidateTransformFinite(deltaY, nameof(deltaY));
        PdfAnnotation annotation = GetTransformTarget(pdf, objectNumber, readOptions);
        var rectangle = new PdfPageRectangle(
            annotation.X1 + deltaX,
            annotation.Y1 + deltaY,
            annotation.X2 + deltaX,
            annotation.Y2 + deltaY);
        return TransformAnnotation(pdf, annotation, rectangle, readOptions);
    }

    internal static PdfAnnotationEditResult ResizeAnnotation(
        byte[] pdf,
        int objectNumber,
        PdfPageRectangle rectangle,
        PdfLoadOptions? readOptions) {
        Guard.NotNull(rectangle, nameof(rectangle));
        PdfAnnotation annotation = GetTransformTarget(pdf, objectNumber, readOptions);
        return TransformAnnotation(pdf, annotation, rectangle, readOptions);
    }

    private static PdfAnnotationEditResult TransformAnnotation(
        byte[] pdf,
        PdfAnnotation annotation,
        PdfPageRectangle target,
        PdfLoadOptions? readOptions) {
        if (annotation.ObjectNumber is not int objectNumber) {
            throw new NotSupportedException("Direct annotation dictionaries cannot be transformed safely.");
        }
        if (annotation.CalloutLine.Count > 0) {
            throw new NotSupportedException("Free-text callout geometry cannot yet be transformed safely.");
        }

        double sourceWidth = annotation.Width;
        double sourceHeight = annotation.Height;
        if (sourceWidth <= 0D || sourceHeight <= 0D) {
            throw new NotSupportedException("Annotations with an empty source rectangle cannot be transformed safely.");
        }

        double scaleX = target.Width / sourceWidth;
        double scaleY = target.Height / sourceHeight;
        bool preserveAuthoredAppearance = annotation.HasNormalAppearance;
        bool canGenerateMissingAppearance =
            !preserveAuthoredAppearance &&
            PdfAnnotationFlattener.IsSupportedVisualAnnotation(annotation.Subtype);
        var options = new PdfAnnotationUpdateOptions {
            Rectangle = new[] { target.Left, target.Bottom, target.Right, target.Top },
            QuadPoints = TransformPairs(annotation.QuadPoints, annotation, target, scaleX, scaleY),
            Vertices = TransformPairs(annotation.Vertices, annotation, target, scaleX, scaleY),
            Line = TransformPairs(annotation.LineCoordinates, annotation, target, scaleX, scaleY),
            InkPaths = annotation.InkList.Count == 0
                ? null
                : annotation.InkList.Select(path => (IReadOnlyList<double>)TransformPairs(path, annotation, target, scaleX, scaleY)!).ToArray(),
            RegenerateAppearance = canGenerateMissingAppearance,
            PreserveAppearance = preserveAuthoredAppearance || !canGenerateMissingAppearance
        };
        return UpdateAnnotation(pdf, objectNumber, options, readOptions);
    }

    private static double[]? TransformPairs(
        IReadOnlyList<double> coordinates,
        PdfAnnotation source,
        PdfPageRectangle target,
        double scaleX,
        double scaleY) {
        if (coordinates.Count == 0) return null;
        var transformed = new double[coordinates.Count];
        for (int i = 0; i < coordinates.Count; i += 2) {
            transformed[i] = target.Left + ((coordinates[i] - source.X1) * scaleX);
            transformed[i + 1] = target.Bottom + ((coordinates[i + 1] - source.Y1) * scaleY);
        }
        return transformed;
    }

    private static PdfAnnotation GetTransformTarget(byte[] pdf, int objectNumber, PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        if (objectNumber <= 0) throw new ArgumentOutOfRangeException(nameof(objectNumber), "Annotation object number must be positive.");
        return PdfInspector.Inspect(pdf, readOptions).Annotations.SingleOrDefault(annotation => annotation.ObjectNumber == objectNumber)
            ?? throw new ArgumentException("PDF annotation object was not found: " + objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".", nameof(objectNumber));
    }

    private static void ValidateTransformFinite(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "Annotation movement offset must be finite.");
        }
    }
}
