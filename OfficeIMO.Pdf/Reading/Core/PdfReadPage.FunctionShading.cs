using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private bool TryReadFunctionShading(PdfDictionary dictionary, OfficeIccRenderingIntent intent,
        PageContentBudget budget, out PdfPageShadingResource shading) {
        shading = default;
        double[] domain = { 0D, 1D, 0D, 1D };
        if (dictionary.Items.TryGetValue("Domain", out var domainObject) && ResolveObject(domainObject) is not PdfNull &&
            !TryReadExactFiniteNumberArray(domainObject, 4, out domain)) return false;
        if (domain[0] >= domain[1] || domain[2] >= domain[3]) return false;
        double[] matrix = { 1D, 0D, 0D, 1D, 0D, 0D };
        if (dictionary.Items.TryGetValue("Matrix", out var matrixObject) && ResolveObject(matrixObject) is not PdfNull &&
            !TryReadExactFiniteNumberArray(matrixObject, 6, out matrix)) return false;
        var transform = new OfficeTransform(matrix[0], matrix[1], matrix[2], matrix[3], matrix[4], matrix[5]);
        if (!transform.TryInvert(out _)) return false;
        double[]? bounds = null;
        if (dictionary.Items.TryGetValue("BBox", out var boundsObject) && ResolveObject(boundsObject) is not PdfNull) {
            if (!TryReadExactFiniteNumberArray(boundsObject, 4, out var parsedBounds) ||
                parsedBounds[0] >= parsedBounds[2] || parsedBounds[1] >= parsedBounds[3]) return false;
            bounds = parsedBounds;
        }
        if (!dictionary.Items.TryGetValue("ColorSpace", out var colorSpaceObject) ||
            !TryReadColorSpaceResource(colorSpaceObject, budget.TryConsumeColorFunctionEvaluation,
                budget.ColorFunctionResolutionContext, out var colorSpace) || colorSpace.ComponentCount < 1 ||
            !dictionary.Items.TryGetValue("Function", out var functionObject)) return false;
        PdfArray? array = ResolveArray(functionObject);
        if (array != null && array.Items.Count != colorSpace.ComponentCount) return false;
        var functions = new PdfColorFunction[array == null ? 1 : array.Items.Count];
        for (int index = 0; index < functions.Length; index++) {
            if (!PdfColorSpaceFunctionResolver.TryCreateFunction(array == null ? functionObject : array.Items[index],
                2, array == null ? colorSpace.ComponentCount : 1, _objects, _limits.MaxDecodedStreamBytes,
                budget.ColorFunctionResolutionContext, out var function)) return false;
            for (int dimension = 0; dimension < 2; dimension++) {
                if (function.Domain[dimension * 2] > domain[dimension * 2] ||
                    function.Domain[dimension * 2 + 1] < domain[dimension * 2 + 1]) return false;
            }
            functions[index] = function;
        }
        shading = new PdfPageShadingResource(new PdfFunctionShading(functions, colorSpace, domain, transform, bounds,
            intent, EffectiveOutputIntentColorTransform, budget.ChargeFunctionShadingWork, budget.CancellationToken,
            budget.FunctionShadingScale, budget.ChargeFunctionShadingPixels, dictionary));
        return true;
    }
}
