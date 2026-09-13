using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class PdfHtmlConverterExtensions {
    private static void AppendPositionedNativeText(StringBuilder builder, PdfCore.PdfLogicalPage page,
        PositionedPageGeometry geometry, PdfToHtmlOptions options) {
        string Number(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);
        builder.Append("<svg class=\"pdf-native-text\" xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 ");
        builder.Append(Number(geometry.Width));
        builder.Append(' ');
        builder.Append(Number(geometry.Height));
        builder.Append("\" style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\">");
        var emitted = new HashSet<PdfCore.PdfTextSpan>();
        foreach (var block in page.TextBlocks) {
            builder.Append(block.Kind == PdfCore.PdfLogicalElementKind.Heading
                ? "<g role=\"heading\" aria-level=\"1\">" : "<g>");
            foreach (var span in block.Spans) {
                options.CancellationToken.ThrowIfCancellationRequested();
                if (!emitted.Add(span) || !span.IsVisible || string.IsNullOrEmpty(span.Text)) continue;
                PositionedPoint point = geometry.TransformPoint(span.X, span.Y);
                builder.Append("<text x=\"").Append(Number(point.Left)).Append("\" y=\"").Append(Number(point.Top));
                builder.Append("\" font-family=\"Arial, sans-serif\" font-size=\"").Append(Number(span.FontSize));
                builder.Append("\" font-weight=\"").Append(span.IsBold ? "700" : "400");
                builder.Append("\" font-style=\"").Append(span.IsItalic ? "italic" : "normal").Append('"');
                if (span.Color is { } color) {
                    builder.Append(" fill=\"rgb(").Append(color.R).Append(',').Append(color.G).Append(',').Append(color.B).Append(")\"");
                }
                if (span.Advance > 0) builder.Append(" textLength=\"").Append(Number(span.Advance)).Append("\" lengthAdjust=\"spacingAndGlyphs\"");
                double rotation = geometry.RotationDegrees - span.RotationDegrees;
                if (rotation != 0) builder.Append(" transform=\"rotate(").Append(Number(rotation)).Append(' ')
                    .Append(Number(point.Left)).Append(' ').Append(Number(point.Top)).Append(")\"");
                builder.Append(" xml:space=\"preserve\">");
                AppendHtmlText(builder, span.Text);
                builder.AppendLine("</text>");
            }
            builder.AppendLine("</g>");
        }
        builder.AppendLine("</svg>");
        AddWarning(options, "PositionedFontSubstitution",
            "Positioned HTML preserves source text placement and size using a browser font. Original PDF fonts are not embedded; letter shapes may differ.",
            PdfCore.PdfConversionWarningSeverity.Warning);
    }

    private static void AppendPositionedAppearanceTextLayer(StringBuilder builder, PdfCore.PdfLogicalPage page,
        PositionedPageGeometry geometry, PdfToHtmlOptions options) {
        string Number(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);
        builder.Append("<svg class=\"pdf-text-overlay\" fill=\"transparent\" xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 ");
        builder.Append(Number(geometry.Width)).Append(' ').Append(Number(geometry.Height));
        builder.AppendLine("\" style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\">");
        foreach (PdfCore.PdfLogicalTextBlock block in page.TextBlocks) {
            options.CancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfTextSpan? sourceSpan = block.Spans.FirstOrDefault(span => span.IsVisible && !string.IsNullOrEmpty(span.Text));
            if (sourceSpan is null || string.IsNullOrEmpty(block.Text)) continue;
            PositionedPoint point = geometry.TransformPoint(block.XStart, block.BaselineY);
            builder.Append(block.Kind == PdfCore.PdfLogicalElementKind.Heading
                ? "<text role=\"heading\" aria-level=\"1\"" : "<text");
            builder.Append(" x=\"").Append(Number(point.Left)).Append("\" y=\"").Append(Number(point.Top));
            builder.Append("\" font-family=\"Arial, sans-serif\" font-size=\"").Append(Number(block.FontSize));
            builder.Append("\" font-weight=\"").Append(sourceSpan.IsBold ? "700" : "400");
            builder.Append("\" font-style=\"").Append(sourceSpan.IsItalic ? "italic" : "normal").Append('"');
            double width = block.XEnd - block.XStart;
            if (width > 0D) builder.Append(" textLength=\"").Append(Number(width)).Append("\" lengthAdjust=\"spacingAndGlyphs\"");
            double rotation = geometry.RotationDegrees - sourceSpan.RotationDegrees;
            if (rotation != 0D) builder.Append(" transform=\"rotate(").Append(Number(rotation)).Append(' ')
                .Append(Number(point.Left)).Append(' ').Append(Number(point.Top)).Append(")\"");
            builder.Append(" xml:space=\"preserve\">");
            AppendHtmlText(builder, block.Text);
            builder.AppendLine("</text>");
        }
        builder.AppendLine("</svg>");
    }
}
