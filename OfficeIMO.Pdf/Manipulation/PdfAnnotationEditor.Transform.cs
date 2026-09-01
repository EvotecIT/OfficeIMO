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
        LineAuxiliaryGeometry lineAuxiliary = string.Equals(annotation.Subtype, "Line", StringComparison.OrdinalIgnoreCase)
            ? ReadLineAuxiliaryGeometry(pdf, objectNumber, readOptions)
            : default;
        double lineNormalScale = lineAuxiliary == default
            ? 1D
            : GetLineNormalScale(annotation.LineCoordinates, scaleX, scaleY);
        var options = new PdfAnnotationUpdateOptions {
            Rectangle = new[] { target.Left, target.Bottom, target.Right, target.Top },
            RectangleDifferences = TransformRectangleDifferences(annotation.RectangleDifferences, scaleX, scaleY),
            QuadPoints = TransformPairs(annotation.QuadPoints, annotation, target, scaleX, scaleY),
            Vertices = TransformPairs(annotation.Vertices, annotation, target, scaleX, scaleY),
            Line = TransformPairs(annotation.LineCoordinates, annotation, target, scaleX, scaleY),
            LineLeaderLength = lineAuxiliary.LeaderLength.HasValue
                ? lineAuxiliary.LeaderLength.Value * lineNormalScale
                : null,
            LineLeaderExtension = lineAuxiliary.LeaderExtension.HasValue
                ? lineAuxiliary.LeaderExtension.Value * lineNormalScale
                : null,
            LineCaptionOffset = lineAuxiliary.CaptionOffset is null
                ? null
                : [lineAuxiliary.CaptionOffset[0] * scaleX, lineAuxiliary.CaptionOffset[1] * scaleY],
            InkPaths = annotation.InkList.Count == 0
                ? null
                : annotation.InkList.Select(path => (IReadOnlyList<double>)TransformPairs(path, annotation, target, scaleX, scaleY)!).ToArray(),
            RegenerateAppearance = canGenerateMissingAppearance,
            PreserveAppearance = preserveAuthoredAppearance || !canGenerateMissingAppearance
        };
        return UpdateAnnotation(pdf, objectNumber, options, readOptions);
    }

    private static double[]? TransformRectangleDifferences(
        IReadOnlyList<double> rectangleDifferences,
        double scaleX,
        double scaleY) =>
        rectangleDifferences.Count == 0
            ? null
            : [
                rectangleDifferences[0] * scaleX,
                rectangleDifferences[1] * scaleY,
                rectangleDifferences[2] * scaleX,
                rectangleDifferences[3] * scaleY
            ];

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

    private static double GetLineNormalScale(IReadOnlyList<double> coordinates, double scaleX, double scaleY) {
        if (coordinates.Count != 4) return 1D;
        double dx = coordinates[2] - coordinates[0];
        double dy = coordinates[3] - coordinates[1];
        double length = Math.Sqrt(dx * dx + dy * dy);
        if (length <= 0D) return 1D;
        double normalX = -dy / length;
        double normalY = dx / length;
        return Math.Sqrt(normalX * normalX * scaleX * scaleX + normalY * normalY * scaleY * scaleY);
    }

    private static LineAuxiliaryGeometry ReadLineAuxiliaryGeometry(
        byte[] pdf,
        int objectNumber,
        PdfLoadOptions? readOptions) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf, readOptions);
        if (!objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect) ||
            indirect.Value is not PdfDictionary dictionary) {
            return default;
        }

        double? leaderLength = ReadOptionalNumber(objects, dictionary, "LL");
        double? leaderExtension = ReadOptionalNumber(objects, dictionary, "LLE");
        double[]? captionOffset = ReadOptionalNumberPair(objects, dictionary, "CO");
        return new LineAuxiliaryGeometry(leaderLength, leaderExtension, captionOffset);
    }

    private static double? ReadOptionalNumber(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        string key) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value)) return null;
        if (PdfObjectLookup.Resolve(objects, value) is not PdfNumber number) {
            throw new NotSupportedException("Line annotation /" + key + " geometry is malformed and cannot be transformed safely.");
        }
        return number.Value;
    }

    private static double[]? ReadOptionalNumberPair(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        string key) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value)) return null;
        if (PdfObjectLookup.Resolve(objects, value) is not PdfArray array ||
            array.Items.Count != 2 ||
            PdfObjectLookup.Resolve(objects, array.Items[0]) is not PdfNumber x ||
            PdfObjectLookup.Resolve(objects, array.Items[1]) is not PdfNumber y) {
            throw new NotSupportedException("Line annotation /" + key + " geometry is malformed and cannot be transformed safely.");
        }
        return [x.Value, y.Value];
    }

    private readonly record struct LineAuxiliaryGeometry(
        double? LeaderLength,
        double? LeaderExtension,
        double[]? CaptionOffset);

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
