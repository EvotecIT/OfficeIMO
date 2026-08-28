using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfPrintProductionColorInspector {
    private static bool IsStructurallyInspectableShading(
        ShadingContext context,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes) {
        PdfDictionary shading = context.Dictionary;
        if (componentCount < 1 ||
            !TryResolveInteger(shading, "ShadingType", objects, maximumObjectDepth, 1, 7, out int shadingType) ||
            !HasOptionalExactFiniteNumberArray(shading, "BBox", 4, objects, maximumObjectDepth) ||
            !HasOptionalBoolean(shading, "AntiAlias", objects, maximumObjectDepth)) {
            return false;
        }

        bool functionRequired = shadingType <= 3;
        bool hasFunction = shading.Items.TryGetValue("Function", out PdfObject? functionObject) &&
            ResolveObject(objects, functionObject, 0, maximumObjectDepth) is not PdfNull;
        if (functionRequired && !hasFunction) return false;
        if (hasFunction) {
            bool validFunction = shadingType == 1
                ? PdfColorSpaceFunctionResolver.TryCreateFunction(
                    functionObject, 2, componentCount, objects, maximumDecodedStreamBytes, out _)
                : PdfColorSpaceFunctionResolver.TryCreateShadingFunction(
                    functionObject, componentCount, objects, maximumDecodedStreamBytes, out _);
            if (!validFunction) return false;
        }

        switch (shadingType) {
            case 1:
                return context.Stream == null &&
                    HasOptionalExactFiniteNumberArray(shading, "Domain", 4, objects, maximumObjectDepth) &&
                    HasOptionalExactFiniteNumberArray(shading, "Matrix", 6, objects, maximumObjectDepth);
            case 2:
                return context.Stream == null &&
                    HasExactFiniteNumberArray(shading, "Coords", 4, objects, maximumObjectDepth, out _) &&
                    HasOptionalExactFiniteNumberArray(shading, "Domain", 2, objects, maximumObjectDepth) &&
                    HasOptionalBooleanArray(shading, "Extend", 2, objects, maximumObjectDepth);
            case 3:
                return context.Stream == null &&
                    HasExactFiniteNumberArray(shading, "Coords", 6, objects, maximumObjectDepth, out double[] radialCoords) &&
                    radialCoords[2] >= 0D && radialCoords[5] >= 0D &&
                    HasOptionalExactFiniteNumberArray(shading, "Domain", 2, objects, maximumObjectDepth) &&
                    HasOptionalBooleanArray(shading, "Extend", 2, objects, maximumObjectDepth);
            default:
                return IsStructurallyInspectableMeshShading(
                    context,
                    shadingType,
                    componentCount,
                    hasFunction,
                    objects,
                    maximumObjectDepth,
                    maximumDecodedStreamBytes);
        }
    }

    private static bool IsStructurallyInspectableMeshShading(
        ShadingContext context,
        int shadingType,
        int componentCount,
        bool hasFunction,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes) {
        if (context.Stream == null ||
            !TryResolveIntegerFromSet(context.Dictionary, "BitsPerCoordinate", objects, maximumObjectDepth, 1, 2, 4, 8, 12, 16, 24, 32) ||
            !TryResolveIntegerFromSet(context.Dictionary, "BitsPerComponent", objects, maximumObjectDepth, 1, 2, 4, 8, 12, 16) ||
            shadingType != 5 &&
            !TryResolveIntegerFromSet(context.Dictionary, "BitsPerFlag", objects, maximumObjectDepth, 2, 4, 8) ||
            !HasExactFiniteNumberArray(
                context.Dictionary,
                "Decode",
                4 + (2 * (hasFunction ? 1 : componentCount)),
                objects,
                maximumObjectDepth,
                out _) ||
            shadingType == 5 &&
            !TryResolveInteger(context.Dictionary, "VerticesPerRow", objects, maximumObjectDepth, 2, int.MaxValue, out _) ||
            !StreamDecoder.TryDecode(
                context.Stream.Dictionary,
                context.Stream.Data,
                maximumDecodedStreamBytes,
                out byte[] decoded,
                objects) ||
            decoded.Length == 0) {
            return false;
        }

        return true;
    }

    private static bool HasExactFiniteNumberArray(
        PdfDictionary dictionary,
        string key,
        int expectedCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        out double[] values) {
        values = Array.Empty<double>();
        if (!dictionary.Items.TryGetValue(key, out PdfObject? candidate) ||
            ResolveObject(objects, candidate, 0, maximumObjectDepth) is not PdfArray array ||
            array.Items.Count != expectedCount) return false;
        values = new double[expectedCount];
        for (int index = 0; index < expectedCount; index++) {
            if (ResolveObject(objects, array.Items[index], 0, maximumObjectDepth) is not PdfNumber number ||
                double.IsNaN(number.Value) || double.IsInfinity(number.Value)) {
                values = Array.Empty<double>();
                return false;
            }
            values[index] = number.Value;
        }
        return true;
    }

    private static bool HasOptionalExactFiniteNumberArray(
        PdfDictionary dictionary,
        string key,
        int expectedCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) =>
        !dictionary.Items.TryGetValue(key, out PdfObject? candidate) ||
        ResolveObject(objects, candidate, 0, maximumObjectDepth) is PdfNull ||
        HasExactFiniteNumberArray(dictionary, key, expectedCount, objects, maximumObjectDepth, out _);

    private static bool HasOptionalBooleanArray(
        PdfDictionary dictionary,
        string key,
        int expectedCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? candidate) ||
            ResolveObject(objects, candidate, 0, maximumObjectDepth) is PdfNull) return true;
        if (ResolveObject(objects, candidate, 0, maximumObjectDepth) is not PdfArray array ||
            array.Items.Count != expectedCount) return false;
        return array.Items.All(item =>
            ResolveObject(objects, item, 0, maximumObjectDepth) is PdfBoolean);
    }

    private static bool HasOptionalBoolean(
        PdfDictionary dictionary,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) =>
        !dictionary.Items.TryGetValue(key, out PdfObject? candidate) ||
        ResolveObject(objects, candidate, 0, maximumObjectDepth) is PdfNull or PdfBoolean;

    private static bool TryResolveInteger(
        PdfDictionary dictionary,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int minimum,
        int maximum,
        out int value) {
        value = 0;
        if (!TryResolveNumber(dictionary, key, objects, maximumObjectDepth, out double number) ||
            double.IsNaN(number) || double.IsInfinity(number) || number != Math.Truncate(number) ||
            number < minimum || number > maximum) return false;
        value = (int)number;
        return true;
    }

    private static bool TryResolveIntegerFromSet(
        PdfDictionary dictionary,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        params int[] allowed) =>
        TryResolveInteger(dictionary, key, objects, maximumObjectDepth, int.MinValue, int.MaxValue, out int value) &&
        allowed.Contains(value);
}
