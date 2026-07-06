namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private void ReadAnnotationPathGeometryMetadata(
        PdfDictionary annotation,
        out IReadOnlyList<double> quadPoints,
        out IReadOnlyList<double> lineCoordinates,
        out IReadOnlyList<double> vertices,
        out IReadOnlyList<IReadOnlyList<double>> inkList) {
        quadPoints = TryReadEvenNumberArray(annotation, "QuadPoints", minimumCount: 8, requiredMultiple: 8) ?? Array.Empty<double>();
        lineCoordinates = TryReadEvenNumberArray(annotation, "L", minimumCount: 4, requiredMultiple: 4) ?? Array.Empty<double>();
        vertices = TryReadEvenNumberArray(annotation, "Vertices", minimumCount: 4, requiredMultiple: 2) ?? Array.Empty<double>();
        System.Collections.ObjectModel.ReadOnlyCollection<IReadOnlyList<double>>? parsedInkList = TryReadInkList(annotation);
        inkList = parsedInkList != null
            ? parsedInkList
            : Array.Empty<IReadOnlyList<double>>();
    }

    private double[]? TryReadEvenNumberArray(PdfDictionary dictionary, string key, int minimumCount, int requiredMultiple) {
        PdfArray? array = ResolveArray(dictionary.Items.TryGetValue(key, out PdfObject? value) ? value : null);
        if (array == null ||
            array.Items.Count < minimumCount ||
            array.Items.Count % requiredMultiple != 0) {
            return null;
        }

        var values = new double[array.Items.Count];
        for (int i = 0; i < array.Items.Count; i++) {
            if (ResolveObject(array.Items[i]) is not PdfNumber number ||
                !IsFinite(number.Value)) {
                return null;
            }

            values[i] = number.Value;
        }

        return values;
    }

    private System.Collections.ObjectModel.ReadOnlyCollection<IReadOnlyList<double>>? TryReadInkList(PdfDictionary annotation) {
        PdfArray? inkList = ResolveArray(annotation.Items.TryGetValue("InkList", out PdfObject? inkListObject) ? inkListObject : null);
        if (inkList == null || inkList.Items.Count == 0) {
            return null;
        }

        var paths = new List<IReadOnlyList<double>>();
        for (int i = 0; i < inkList.Items.Count; i++) {
            PdfArray? path = ResolveArray(inkList.Items[i]);
            if (path == null ||
                path.Items.Count < 4 ||
                path.Items.Count % 2 != 0) {
                return null;
            }

            var coordinates = new double[path.Items.Count];
            for (int coordinateIndex = 0; coordinateIndex < path.Items.Count; coordinateIndex++) {
                if (ResolveObject(path.Items[coordinateIndex]) is not PdfNumber coordinate ||
                    !IsFinite(coordinate.Value)) {
                    return null;
                }

                coordinates[coordinateIndex] = coordinate.Value;
            }

            paths.Add(coordinates);
        }

        return paths.Count == 0 ? null : paths.AsReadOnly();
    }
}
