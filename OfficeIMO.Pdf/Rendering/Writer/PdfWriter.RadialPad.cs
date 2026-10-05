using System.Collections.Generic;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static string GradientPaintFormName(string shadingName) => "GR_" + shadingName;

    private static void PaintGradient(ContentStreamBuilder content, OfficeShape shape, string shadingName) {
        if (shape.FillRadialGradient?.OutsideColor != null) content.XObject(GradientPaintFormName(shadingName));
        else content.Shading(shadingName);
    }

    // Compose endpoint background and opaque RGB shading before applying the
    // caller's alpha mask/opacity once. This avoids cone-edge coverage seams and
    // does not depend on a consumer implementing shading /Background.
    private static void AddRadialPadResources(IList<byte[]> objects, PageShading shading, int shadingId,
        List<(string Name, int Id)> xobjects, PdfPrintColorTransform? printColorTransform, CancellationToken cancellationToken) {
        if (shading.OutsideColor is not OfficeColor outside) return;
        var buffer = new StringBuilder();
        new ContentStreamBuilder(buffer).FillColor(PdfColor.FromOfficeColor(outside))
            .Rectangle(shading.AlphaLeft, shading.AlphaBottom, shading.AlphaRight - shading.AlphaLeft, shading.AlphaTop - shading.AlphaBottom)
            .FillPath().Shading("A");
        string content = buffer.ToString();
        if (printColorTransform != null) content = printColorTransform.NormalizeGeneratedContent(content, cancellationToken);
        string entries = "/Type /XObject /Subtype /Form /FormType 1 /BBox [" +
            F(shading.AlphaLeft) + " " + F(shading.AlphaBottom) + " " + F(shading.AlphaRight) + " " + F(shading.AlphaTop) + "]" +
            " /Group << /S /Transparency /CS " + (printColorTransform == null ? "/DeviceRGB" : "/DeviceCMYK") + " /I true >>" +
            " /Resources << /Shading << /A " + shadingId + " 0 R >> >>";
        int formId = AddFlateStreamObject(objects, Encoding.ASCII.GetBytes(content), entries);
        xobjects.Add(("/" + GradientPaintFormName(shading.Name), formId));
    }
}
