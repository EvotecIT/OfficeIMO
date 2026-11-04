using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

internal static partial class PowerPointSlideImageRenderer {
    private static void ReportTextBoxOverflow(OfficeDrawingElement element, PowerPointTextBox textBox,
        List<OfficeImageExportDiagnostic> diagnostics, Func<string?, double, string?, OfficeFontStyle, double> measure) {
        bool clipped = false;
        if (element is OfficeDrawingRichText rich) {
            double availableHeight = Math.Max(0D, rich.Height - rich.Padding.Vertical);
            var layout = OfficeDrawingTextLayout.Create(rich, rich.Width - rich.Padding.Horizontal,
                availableHeight, measure);
            clipped = layout.Clipped || OfficeTextPlacement.ResolveTop(0D, availableHeight, layout.Height, rich.VerticalAlignment) +
                Math.Max(layout.Height, OfficeDrawingTextLayout.PaintedHeight(layout)) > availableHeight + 0.001D;
        } else if (element is OfficeDrawingText plain) {
            double size = plain.Font.Size;
            double lineHeightFactor = OfficeDrawingTextLayout.ResolveLineHeightFactor(plain.LineHeight, size);
            double availableHeight = Math.Max(0D, plain.Height - plain.Padding.Vertical);
            var layout = OfficeTextLayoutEngine.LayoutTextBlock(plain.Text, size,
                plain.Width - plain.Padding.Horizontal, plain.Height - plain.Padding.Vertical,
                lineHeightFactor, Math.Min(6D, size),
                (value, fontSize) => measure(value, fontSize, plain.Font.FamilyName, plain.Font.Style),
                plain.WrapText, shrinkToFit: plain.ShrinkToFit, paragraphIndent: plain.ParagraphIndent);
            clipped = layout.Clipped || OfficeTextPlacement.ResolveTop(0D, availableHeight, layout.Height, plain.VerticalAlignment) +
                Math.Max(layout.Height, OfficeDrawingTextLayout.PaintedHeight(layout)) > availableHeight + 0.001D;
        }
        if (clipped) diagnostics.Add(new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Warning, global::OfficeIMO.OfficeDiagnosticCodes.PowerpointTextOverflow,
            "Clipped PowerPoint text because it exceeds the available text frame.",
            DescribeShape(textBox), OfficeConversionLossKind.Omission));
    }
}
