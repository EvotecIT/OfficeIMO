using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static int AddRadialSpreadShading(IList<byte[]> objects, PageShading shading, bool alphaOnly) {
        PdfVisualResourceDictionaryBuilder.ValidateStops(shading.Stops);
        string domain = PdfNumberFormatter.Precise(shading.AlphaLeft) + " " + PdfNumberFormatter.Precise(shading.AlphaRight) + " " +
            PdfNumberFormatter.Precise(shading.AlphaBottom) + " " + PdfNumberFormatter.Precise(shading.AlphaTop);
        // Large RGB trees use separate component functions so portable PDF
        // interpreters need neither deep procedures nor oversized branch jumps.
        bool separateComponents = !alphaOnly && shading.Stops.Count > 32;
        var functionIds = new List<int>();
        for (int channel = 0; channel < (separateComponents ? 3 : 1); channel++) {
            string program = PdfRadialSpreadFunction.Build(shading.X0, shading.Y0, shading.R0,
                shading.X1, shading.Y1, shading.R1, shading.Stops, shading.SpreadMode, shading.ColorInterpolation,
                shading.OutsideColor, shading.R1 == 0D && shading.OutsideColor.HasValue, alphaOnly,
                separateComponents ? channel : -1);
            string range = alphaOnly || separateComponents ? "0 1" : "0 1 0 1 0 1";
            functionIds.Add(AddFlateStreamObject(objects, Encoding.ASCII.GetBytes(program),
                "/FunctionType 4 /Domain [" + domain + "] /Range [" + range + "]"));
        }
        string function = separateComponents ? "[" + string.Join(" ", functionIds.Select(id => id + " 0 R")) + "]"
            : functionIds[0] + " 0 R";
        string colorSpace = alphaOnly ? "/DeviceGray" : shading.ColorInterpolation == OfficeGradientColorInterpolation.LinearRgb
            ? PdfVisualResourceDictionaryBuilder.LinearRgbColorSpace : "/DeviceRGB";
        return AddObject(objects, "<< /ShadingType 1 /ColorSpace " + colorSpace + " /Domain [" + domain +
            "] /Function " + function + " >>\n");
    }
}
