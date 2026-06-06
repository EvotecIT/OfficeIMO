using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class PdfHtmlConverter {
    private static void AppendPositionedPage(StringBuilder builder, PdfCore.PdfLogicalPage page, PdfHtmlSaveOptions options) {
        builder.Append("<section class=\"pdf-page\" data-page-number=\"");
        builder.Append(page.PageNumber.ToString(CultureInfo.InvariantCulture));
        builder.Append("\" style=\"width:");
        builder.Append(Points(page.Width));
        builder.Append(";height:");
        builder.Append(Points(page.Height));
        builder.AppendLine(";\">");

        for (int i = 0; i < page.TextBlocks.Count; i++) {
            PdfCore.PdfLogicalTextBlock block = page.TextBlocks[i];
            string cssClass = block.Kind == PdfCore.PdfLogicalElementKind.Heading
                ? "pdf-text pdf-heading"
                : block.Kind == PdfCore.PdfLogicalElementKind.ListItem
                    ? "pdf-text pdf-list-item"
                    : "pdf-text";
            builder.Append("<div class=\"");
            builder.Append(cssClass);
            builder.Append("\" style=\"left:");
            builder.Append(Points(block.XStart));
            builder.Append(";top:");
            builder.Append(Points(Math.Max(0D, page.Height - block.BaselineY)));
            builder.Append(";width:");
            builder.Append(Points(Math.Max(1D, block.XEnd - block.XStart)));
            builder.Append(";\">");
            builder.Append(HtmlText(block.Text));
            builder.AppendLine("</div>");
        }

        for (int i = 0; i < page.Tables.Count; i++) {
            AppendPositionedTable(builder, page, page.Tables[i]);
        }

        if (options.IncludeImagePlaceholders) {
            AppendPositionedImagePlaceholders(builder, page, options);
        }

        if (options.IncludeLinkAnnotations) {
            for (int i = 0; i < page.Links.Count; i++) {
                AppendPositionedLink(builder, page, page.Links[i]);
            }
        }

        if (options.IncludeFormWidgets) {
            for (int i = 0; i < page.FormWidgets.Count; i++) {
                AppendPositionedFormWidget(builder, page, page.FormWidgets[i]);
            }
        }

        builder.AppendLine("</section>");
    }

    private static void AppendPositionedTable(StringBuilder builder, PdfCore.PdfLogicalPage page, PdfCore.PdfLogicalTable table) {
        if (table.Rows.Count == 0) {
            return;
        }

        double left = table.Columns.Count > 0 ? table.Columns[0].From : 0D;
        double width = table.Columns.Count > 0 ? Math.Max(1D, table.Columns[table.Columns.Count - 1].To - left) : 1D;
        double top = Math.Max(0D, page.Height - Math.Max(table.YTop, table.YBottom));
        double height = Math.Max(1D, Math.Abs(table.YTop - table.YBottom));

        builder.Append("<table class=\"pdf-table\" data-detection-kind=\"");
        builder.Append(HtmlAttribute(table.DetectionKind));
        builder.Append("\" style=\"left:");
        builder.Append(Points(left));
        builder.Append(";top:");
        builder.Append(Points(top));
        builder.Append(";width:");
        builder.Append(Points(width));
        builder.Append(";height:");
        builder.Append(Points(height));
        builder.AppendLine(";\">");
        AppendTableRows(builder, table);
        builder.AppendLine("</table>");
    }

    private static void AppendPositionedLink(StringBuilder builder, PdfCore.PdfLogicalPage page, PdfCore.PdfLogicalLinkAnnotation link) {
        string target = link.Uri ?? link.DestinationName ?? string.Empty;
        if (target.Length == 0) {
            return;
        }

        builder.Append("<a class=\"pdf-link\" style=\"left:");
        builder.Append(Points(link.X1));
        builder.Append(";top:");
        builder.Append(Points(Math.Max(0D, page.Height - link.Y2)));
        builder.Append(";width:");
        builder.Append(Points(Math.Max(1D, link.Width)));
        builder.Append(";height:");
        builder.Append(Points(Math.Max(1D, link.Height)));
        builder.Append("\"");
        if (link.Uri is not null && IsSafeLinkUri(link.Uri)) {
            builder.Append(" href=\"");
            builder.Append(HtmlAttribute(link.Uri));
            builder.Append('"');
        } else if (link.Uri is not null) {
            builder.Append(" data-unsafe-href=\"");
            builder.Append(HtmlAttribute(link.Uri));
            builder.Append('"');
        } else {
            builder.Append(" data-destination=\"");
            builder.Append(HtmlAttribute(target));
            builder.Append('"');
        }

        builder.Append('>');
        builder.Append(HtmlText(!string.IsNullOrWhiteSpace(link.Contents) ? link.Contents! : target));
        builder.AppendLine("</a>");
    }

    private static void AppendPositionedImagePlaceholders(StringBuilder builder, PdfCore.PdfLogicalPage page, PdfHtmlSaveOptions options) {
        if (page.Images.Count == 0) {
            return;
        }

        var unplaced = new List<PdfCore.PdfLogicalImage>();
        for (int imageIndex = 0; imageIndex < page.Images.Count; imageIndex++) {
            PdfCore.PdfLogicalImage image = page.Images[imageIndex];
            if (!image.HasPlacements) {
                unplaced.Add(image);
                continue;
            }

            for (int placementIndex = 0; placementIndex < image.Placements.Count; placementIndex++) {
                AppendPositionedImagePlaceholder(builder, page, image, image.Placements[placementIndex], placementIndex, options);
            }
        }

        if (unplaced.Count == 0) {
            return;
        }

        AddWarning(options, "ImagePlaceholder", "Some images are represented as page-scoped placeholders because no placement invocation was detected.");
        builder.AppendLine("<div class=\"pdf-image-placeholder\" style=\"position:absolute;left:0;bottom:0;\">");
        for (int i = 0; i < unplaced.Count; i++) {
            builder.Append(RenderImageFigure(unplaced[i], options));
        }

        builder.AppendLine("</div>");
    }

    private static void AppendPositionedImagePlaceholder(StringBuilder builder, PdfCore.PdfLogicalPage page, PdfCore.PdfLogicalImage image, PdfCore.PdfImagePlacement placement, int placementIndex, PdfHtmlSaveOptions options) {
        builder.Append("<figure class=\"pdf-image-placeholder\" data-resource=\"");
        builder.Append(HtmlAttribute(image.ResourceName));
        builder.Append("\" data-page-number=\"");
        builder.Append(image.PageNumber.ToString(CultureInfo.InvariantCulture));
        builder.Append("\" data-placement-index=\"");
        builder.Append(placementIndex.ToString(CultureInfo.InvariantCulture));
        builder.Append("\" data-matrix=\"");
        builder.Append(HtmlAttribute(FormatMatrix(placement)));
        builder.Append("\" style=\"position:absolute;left:");
        builder.Append(Points(placement.X));
        builder.Append(";top:");
        builder.Append(Points(Math.Max(0D, page.Height - placement.Y - placement.Height)));
        builder.Append(";width:");
        builder.Append(Points(Math.Max(1D, placement.Width)));
        builder.Append(";height:");
        builder.Append(Points(Math.Max(1D, placement.Height)));
        builder.Append(";\">");
        if (TryBuildEmbeddedImageDataUri(image, options, out string? source)) {
            builder.Append("<img src=\"");
            builder.Append(HtmlAttribute(source!));
            builder.Append("\" alt=\"");
            builder.Append(HtmlAttribute("Image: " + image.ResourceName));
            builder.Append("\" style=\"width:100%;height:100%;object-fit:contain;display:block;\">");
        } else {
            builder.Append("<figcaption>Image: ");
            builder.Append(HtmlText(image.ResourceName));
            builder.Append(" (");
            builder.Append(image.Width.ToString(CultureInfo.InvariantCulture));
            builder.Append('x');
            builder.Append(image.Height.ToString(CultureInfo.InvariantCulture));
            if (!string.IsNullOrWhiteSpace(image.MimeType)) {
                builder.Append(", ");
                builder.Append(HtmlText(image.MimeType!));
            }

            builder.Append(")</figcaption>");
        }

        builder.Append("</figure>");
        builder.AppendLine();
    }

    private static void AppendPositionedFormWidget(StringBuilder builder, PdfCore.PdfLogicalPage page, PdfCore.PdfLogicalFormWidget widget) {
        string name = widget.FieldName ?? widget.FieldType ?? "Field";
        builder.Append("<div class=\"pdf-form-widget\" style=\"left:");
        builder.Append(Points(widget.X1));
        builder.Append(";top:");
        builder.Append(Points(Math.Max(0D, page.Height - widget.Y2)));
        builder.Append(";width:");
        builder.Append(Points(Math.Max(1D, widget.Width)));
        builder.Append(";height:");
        builder.Append(Points(Math.Max(1D, widget.Height)));
        builder.Append(";\">");
        builder.Append(HtmlText(name));
        if (!string.IsNullOrEmpty(widget.Value)) {
            builder.Append(": ");
            builder.Append(HtmlText(widget.Value!));
        }

        builder.AppendLine("</div>");
    }
}
