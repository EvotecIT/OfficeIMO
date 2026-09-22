using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class PdfHtmlConverterExtensions {
    private static void AppendPositionedPage(StringBuilder builder, IReadOnlyList<PdfCore.PdfLogicalPage> pages, int renderIndex, PdfToHtmlOptions options) {
        PdfCore.PdfLogicalPage page = pages[renderIndex];
        PositionedPageGeometry geometry = PositionedPageGeometry.From(page);
        builder.Append("<section class=\"pdf-page\" id=\"");
        builder.Append(GetPageAnchorId(page.PageNumber, pages, renderIndex));
        builder.Append("\" data-page-number=\"");
        builder.Append(page.PageNumber.ToString(CultureInfo.InvariantCulture));
        builder.Append("\" style=\"width:");
        builder.Append(Points(geometry.Width));
        builder.Append(";height:");
        builder.Append(Points(geometry.Height));
        builder.AppendLine(";\">");

        bool hasPageAppearance = TryAppendPageAppearance(builder, page, renderIndex, options);
        bool hasNativeText = page.TextBlocks.Count > 0 && page.TextBlocks.All(block =>
            block.Spans.Count > 0 && block.Spans.All(span => span.IsVisible));
        if (hasPageAppearance && page.TextBlocks.Any(block => block.Spans.Any(span =>
                !string.IsNullOrEmpty(span.Text) && (span.IsVisible || options.IncludeInvisibleTextInAppearanceOverlay)))) {
            // Keep rendered text searchable and selectable. Invisible OCR text is
            // included only when the caller explicitly requests it.
            AppendPositionedAppearanceTextLayer(builder, page, geometry, options);
        }
        if (!hasPageAppearance) {

            // Native PDF spans already carry the placement of table cells. Reflowing a
            // detected table here discards that geometry and can duplicate nearby text.
            // Mixed OCR/native pages still need the logical block/table projection.
            // Only replace that projection when every block can be represented here.
            if (hasNativeText) AppendPositionedNativeText(builder, page, geometry, options);

            for (int i = 0; i < page.TextBlocks.Count; i++) {
                PdfCore.PdfLogicalTextBlock block = page.TextBlocks[i];
                if (hasNativeText) continue;
                if (IsPositionedTextBlockRepresentedByTable(block, page.Tables)) {
                    continue;
                }

                PositionedPoint point = geometry.TransformPoint(block.XStart, block.BaselineY);
                string cssClass = block.Kind == PdfCore.PdfLogicalElementKind.Heading
                    ? "pdf-text pdf-heading"
                    : block.Kind == PdfCore.PdfLogicalElementKind.ListItem
                        ? "pdf-text pdf-list-item"
                        : "pdf-text";
                builder.Append("<div class=\"");
                builder.Append(cssClass);
                builder.Append("\" style=\"left:");
                builder.Append(Points(point.Left));
                builder.Append(";top:");
                builder.Append(Points(Math.Max(0D, point.Top)));
                builder.Append(";width:");
                builder.Append(Points(Math.Max(1D, geometry.ScaleLength(block.XEnd - block.XStart))));
                builder.Append(";font-size:");
                builder.Append(Points(Math.Max(1D, geometry.ScaleLength(block.FontSize > 0D ? block.FontSize : 10D))));
                builder.Append(";\">");
                AppendHtmlText(builder, block.Text);
                builder.AppendLine("</div>");
            }

            if (!hasNativeText) {
                for (int i = 0; i < page.Tables.Count; i++) {
                    AppendPositionedTable(builder, geometry, page.Tables[i]);
                }
            }

            if (page.VectorPrimitiveCount > 0) {
                bool hasOmittedVectors = page.UnrepresentedVectorPrimitiveCount > 0;
                AddWarning(
                    options,
                    "VectorAppearanceNotExported",
                    hasOmittedVectors
                        ? page.UnrepresentedVectorPrimitiveCount.ToString(CultureInfo.InvariantCulture) +
                          " vector primitives were not represented by positioned HTML content or detected table structure."
                        : "PDF vector primitives were reconstructed through detected table structure rather than their exact source appearance.",
                    hasOmittedVectors
                        ? PdfCore.PdfConversionWarningSeverity.Warning
                        : PdfCore.PdfConversionWarningSeverity.Information,
                    hasOmittedVectors
                        ? OfficeConversionLossKind.Omission
                        : OfficeConversionLossKind.Approximation);
            }

            if (options.IncludeImagePlaceholders) {
                AppendPositionedImagePlaceholders(builder, page, page.Images, options);
            }
        }

        if (options.IncludeLinkAnnotations) {
            for (int i = 0; i < page.Links.Count; i++) {
                AppendPositionedLink(builder, geometry, page.Links[i]);
            }
        }

        if (options.IncludeFormWidgets) {
            for (int i = 0; i < page.FormWidgets.Count; i++) {
                AppendPositionedFormWidget(builder, geometry, page.FormWidgets[i]);
            }
        }

        builder.AppendLine("</section>");
    }

    private static bool IsPositionedTextBlockRepresentedByTable(
        PdfCore.PdfLogicalTextBlock block,
        IReadOnlyList<PdfCore.PdfLogicalTable> tables) {
        for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++) {
            if (IsTextBlockRepresentedByTable(block, tables[tableIndex])) {
                return true;
            }
        }

        return false;
    }

    private static void AppendPositionedTable(StringBuilder builder, PositionedPageGeometry geometry, PdfCore.PdfLogicalTable table) {
        if (table.Rows.Count == 0) {
            return;
        }

        double left = table.Columns.Count > 0 ? table.Columns[0].From : 0D;
        double width = table.Columns.Count > 0 ? table.Columns[table.Columns.Count - 1].To - left : 1D;
        double bottom = Math.Min(table.YTop, table.YBottom);
        double height = Math.Abs(table.YTop - table.YBottom);
        PositionedBox box = geometry.TransformBox(left, bottom, width, height);

        builder.Append("<table class=\"pdf-table\" data-detection-kind=\"");
        builder.Append(HtmlAttribute(table.DetectionKind));
        builder.Append("\" style=\"left:");
        builder.Append(Points(box.Left));
        builder.Append(";top:");
        builder.Append(Points(Math.Max(0D, box.Top)));
        builder.Append(";width:");
        builder.Append(Points(Math.Max(1D, box.Width)));
        builder.Append(";height:");
        builder.Append(Points(Math.Max(1D, box.Height)));
        builder.AppendLine(";\">");
        AppendTableRows(builder, table);
        builder.AppendLine("</table>");
    }

    private static void AppendPositionedLink(StringBuilder builder, PositionedPageGeometry geometry, PdfCore.PdfLogicalLinkAnnotation link) {
        if (!HasHtmlLinkTarget(link)) {
            return;
        }

        string label = GetLinkLabel(link);
        PositionedBox box = geometry.TransformBox(link.X1, link.Y1, link.Width, link.Height);
        builder.Append("<a class=\"pdf-link\" style=\"left:");
        builder.Append(Points(box.Left));
        builder.Append(";top:");
        builder.Append(Points(Math.Max(0D, box.Top)));
        builder.Append(";width:");
        builder.Append(Points(Math.Max(1D, box.Width)));
        builder.Append(";height:");
        builder.Append(Points(Math.Max(1D, box.Height)));
        builder.Append("\" aria-label=\"");
        builder.Append(HtmlAttribute(label));
        builder.Append('"');
        AppendLinkTargetAttributes(builder, link);
        builder.AppendLine("></a>");
    }

    internal static void AppendPositionedImagePlaceholders(
        StringBuilder builder,
        PdfCore.PdfLogicalPage page,
        IReadOnlyList<PdfCore.PdfLogicalImage> images,
        PdfToHtmlOptions options) {
        if (images.Count == 0) {
            return;
        }

        for (int imageIndex = 0; imageIndex < images.Count; imageIndex++) {
            PdfCore.PdfLogicalImage image = images[imageIndex];
            if (!image.HasPlacements) {
                ReportHtmlImagePlacementAssessment(
                    image,
                    PdfCore.PdfImagePlacementImportPolicy.Analyze(page, image, placement: null),
                    options);
                continue;
            }

            for (int placementIndex = 0; placementIndex < image.Placements.Count; placementIndex++) {
                AppendPositionedImagePlaceholder(builder, page, image, image.Placements[placementIndex], placementIndex, options);
            }
        }
    }

    private static void AppendPositionedImagePlaceholder(StringBuilder builder, PdfCore.PdfLogicalPage page, PdfCore.PdfLogicalImage image, PdfCore.PdfImagePlacement placement, int placementIndex, PdfToHtmlOptions options) {
        PdfCore.PdfImagePlacementImportAssessment assessment =
            PdfCore.PdfImagePlacementImportPolicy.Analyze(page, image, placement);
        ReportHtmlImagePlacementAssessment(image, assessment, options);
        if (assessment.IsSuppressed) return;
        bool pageRotationUnsupported = assessment.CanImport && page.RotationDegrees % 360 != 0;
        if (pageRotationUnsupported) ReportHtmlImagePageRotationOmission(image, options);

        options.EmittedImagePlaceholderCount++;
        PositionedPageGeometry geometry = PositionedPageGeometry.From(page);
        PositionedBox box = geometry.TransformBox(placement.X, placement.Y, placement.Width, placement.Height);
        builder.Append("<figure class=\"pdf-image-placeholder\" data-resource=\"");
        builder.Append(HtmlAttribute(image.ResourceName));
        builder.Append("\" data-page-number=\"");
        builder.Append(image.PageNumber.ToString(CultureInfo.InvariantCulture));
        builder.Append("\" data-placement-index=\"");
        builder.Append(placementIndex.ToString(CultureInfo.InvariantCulture));
        builder.Append("\" data-matrix=\"");
        builder.Append(HtmlAttribute(FormatMatrix(placement)));
        builder.Append("\" style=\"position:absolute;left:");
        builder.Append(Points(box.Left));
        builder.Append(";top:");
        builder.Append(Points(Math.Max(0D, box.Top)));
        builder.Append(";width:");
        builder.Append(Points(Math.Max(1D, box.Width)));
        builder.Append(";height:");
        builder.Append(Points(Math.Max(1D, box.Height)));
        builder.Append(";\">");
        if (assessment.CanImport && !pageRotationUnsupported &&
            TryBuildEmbeddedImageDataUri(image, options, builder.MaxCapacity - builder.Length, out string? source)) {
            ReportHtmlImagePlacementAssessment(image, assessment, options, imageEmbedded: true);
            builder.Append("<img src=\"");
            builder.Append(HtmlAttribute(source!));
            builder.Append("\" alt=\"");
            builder.Append(HtmlAttribute("Image: " + image.ResourceName));
            builder.Append("\" style=\"width:100%;height:100%;object-fit:contain;display:block;");
            AppendHtmlImageReflectionStyle(builder, placement);
            if (assessment.HasNonDefaultOpacity) {
                builder.Append("opacity:");
                builder.Append(FormatCssOpacity(assessment.Opacity));
                builder.Append(';');
            }
            if (assessment.HasNonNormalBlendMode) {
                builder.Append("mix-blend-mode:");
                builder.Append(ToCssBlendMode(assessment.BlendMode));
                builder.Append(';');
            }
            builder.Append("\">");
        } else {
            builder.Append("<figcaption>Image: ");
            AppendHtmlText(builder, image.ResourceName);
            builder.Append(" (");
            builder.Append(image.Width.ToString(CultureInfo.InvariantCulture));
            builder.Append('x');
            builder.Append(image.Height.ToString(CultureInfo.InvariantCulture));
            if (!string.IsNullOrWhiteSpace(image.MimeType)) {
                builder.Append(", ");
                AppendHtmlText(builder, image.MimeType!);
            }

            builder.Append(")</figcaption>");
        }

        builder.Append("</figure>");
        builder.AppendLine();
    }

    private static void AppendPositionedFormWidget(StringBuilder builder, PositionedPageGeometry geometry, PdfCore.PdfLogicalFormWidget widget) {
        string name = widget.FieldName ?? widget.FieldType ?? "Field";
        PositionedBox box = geometry.TransformBox(widget.X1, widget.Y1, widget.Width, widget.Height);
        builder.Append("<div class=\"pdf-form-widget\" style=\"left:");
        builder.Append(Points(box.Left));
        builder.Append(";top:");
        builder.Append(Points(Math.Max(0D, box.Top)));
        builder.Append(";width:");
        builder.Append(Points(Math.Max(1D, box.Width)));
        builder.Append(";height:");
        builder.Append(Points(Math.Max(1D, box.Height)));
        builder.Append(";\">");
        AppendHtmlText(builder, name);
        if (!string.IsNullOrEmpty(widget.Value)) {
            builder.Append(": ");
            AppendHtmlText(builder, widget.Value!);
        }

        builder.AppendLine("</div>");
    }

    private sealed class PositionedPageGeometry {
        private PositionedPageGeometry(double pageWidth, double pageHeight, int rotationDegrees, double userUnit) {
            Scale = userUnit > 0D && !double.IsNaN(userUnit) && !double.IsInfinity(userUnit) ? userUnit : 1D;
            PageWidth = pageWidth * Scale;
            PageHeight = pageHeight * Scale;
            RotationDegrees = rotationDegrees;
            Width = rotationDegrees == 90 || rotationDegrees == 270 ? PageHeight : PageWidth;
            Height = rotationDegrees == 90 || rotationDegrees == 270 ? PageWidth : PageHeight;
        }

        public double PageWidth { get; }

        public double PageHeight { get; }

        public int RotationDegrees { get; }

        public double Scale { get; }

        public double Width { get; }

        public double Height { get; }

        public static PositionedPageGeometry From(PdfCore.PdfLogicalPage page) {
            int rotation = page.RotationDegrees % 360;
            if (rotation < 0) {
                rotation += 360;
            }

            return new PositionedPageGeometry(page.Width, page.Height, rotation, page.UserUnit.GetValueOrDefault(1D));
        }

        public PositionedPoint TransformPoint(double x, double y) {
            x *= Scale;
            y *= Scale;
            switch (RotationDegrees) {
                case 90:
                    return new PositionedPoint(PageHeight - y, x);
                case 180:
                    return new PositionedPoint(PageWidth - x, PageHeight - y);
                case 270:
                    return new PositionedPoint(y, PageWidth - x);
                default:
                    return new PositionedPoint(x, PageHeight - y);
            }
        }

        public PositionedBox TransformBox(double left, double bottom, double width, double height) {
            left *= Scale;
            bottom *= Scale;
            width = Math.Max(1D, width * Scale);
            height = Math.Max(1D, height * Scale);
            switch (RotationDegrees) {
                case 90:
                    return new PositionedBox(PageHeight - bottom - height, left, height, width);
                case 180:
                    return new PositionedBox(PageWidth - left - width, PageHeight - bottom - height, width, height);
                case 270:
                    return new PositionedBox(bottom, PageWidth - left - width, height, width);
                default:
                    return new PositionedBox(left, PageHeight - bottom - height, width, height);
            }
        }

        public double ScaleLength(double value) => value * Scale;
    }

    private struct PositionedPoint {
        public PositionedPoint(double left, double top) {
            Left = left;
            Top = top;
        }

        public double Left { get; }

        public double Top { get; }
    }

    private struct PositionedBox {
        public PositionedBox(double left, double top, double width, double height) {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }

        public double Left { get; }

        public double Top { get; }

        public double Width { get; }

        public double Height { get; }
    }
}
