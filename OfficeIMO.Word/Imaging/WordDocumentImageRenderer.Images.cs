using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static bool AddImage(WordImage image, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics) {
            WrapTextImage? wrapText = image.WrapText;
            if (wrapText.HasValue && wrapText.Value != WrapTextImage.InLineWithText) {
                if (IsNoWrapAnchoredImage(wrapText.Value)) {
                    return AddNoWrapAnchoredImage(image, context, diagnostics);
                }

                if (wrapText.Value == WrapTextImage.TopAndBottom) {
                    return AddTopAndBottomAnchoredImage(image, context, diagnostics);
                }

                if (IsSideWrappedAnchoredImage(wrapText.Value)) {
                    return AddSideWrappedAnchoredImage(image, wrapText.Value, context, diagnostics);
                }

                AddDiagnostic(
                    diagnostics,
                    "unsupported-word-floating-image",
                    "Skipped a Word image because wrapped image layout requires Word text-wrap semantics that are not implemented by dependency-free export yet.",
                    DescribeImage(image));
                return false;
            }

            return AddInlineImage(image, context, diagnostics);
        }

        private static bool AddInlineImage(WordImage image, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics) {
            if (!TryReadEmbeddedImage(image, diagnostics, out byte[] bytes, out double width, out double height)) {
                return false;
            }

            FitImageToWidth(context.ContentWidth, ref width, ref height);
            double left = context.Left;
            double top = context.Y;
            OfficeImageProjection projection = CreateImageProjection(image, left, top, width, height);
            (double boundsLeft, double boundsTop, double boundsRight, double boundsBottom) = projection.GetDestinationBounds();
            if (boundsLeft < 0D || boundsTop < 0D || boundsRight > context.Drawing.Width || boundsBottom > context.Drawing.Height) {
                AddDiagnostic(
                    diagnostics,
                    "unsupported-word-image",
                    "Skipped a Word image because its inline projection extends outside the current page preview.",
                    DescribeImage(image));
                return false;
            }

            if (!EnsureVerticalSpace(context, boundsBottom - context.Y, diagnostics)) {
                return false;
            }

            context.Drawing.AddImage(bytes, image.ContentType, projection, DescribeImage(image));
            context.Y += boundsBottom - context.Y + ParagraphGapPoints;
            return true;
        }

        private static bool AddNoWrapAnchoredImage(WordImage image, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics) {
            if (!TryReadEmbeddedImage(image, diagnostics, out byte[] bytes, out double width, out double height)) {
                return false;
            }

            if (!TryGetNoWrapAnchorPlacement(image, context, width, height, out double left, out double top)) {
                AddDiagnostic(
                    diagnostics,
                    "unsupported-word-floating-image",
                    "Skipped a Word image because its no-wrap anchor position could not be resolved.",
                    DescribeImage(image));
                return false;
            }

            OfficeImageProjection projection = CreateImageProjection(image, left, top, width, height);
            (double boundsLeft, double boundsTop, double boundsRight, double boundsBottom) = projection.GetDestinationBounds();
            if (boundsLeft < 0D || boundsTop < 0D || boundsRight > context.Drawing.Width || boundsBottom > context.Drawing.Height) {
                AddDiagnostic(
                    diagnostics,
                    "unsupported-word-floating-image",
                    "Skipped a Word image because its no-wrap anchor projects outside the current page preview.",
                    DescribeImage(image));
                return false;
            }

            if (image.WrapText == WrapTextImage.BehindText) {
                context.Drawing.AddImageBehindContent(bytes, image.ContentType, projection, DescribeImage(image));
            } else {
                context.Drawing.AddImage(bytes, image.ContentType, projection, DescribeImage(image));
            }

            return true;
        }

        private static bool AddTopAndBottomAnchoredImage(WordImage image, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics) {
            if (!TryReadEmbeddedImage(image, diagnostics, out byte[] bytes, out double width, out double height)) {
                return false;
            }

            FitImageToWidth(context.ContentWidth, ref width, ref height);
            if (!TryGetNoWrapAnchorPlacement(image, context, width, height, out double left, out double top)) {
                AddDiagnostic(
                    diagnostics,
                    "unsupported-word-floating-image",
                    "Skipped a Word image because its top-and-bottom anchor position could not be resolved.",
                    DescribeImage(image));
                return false;
            }

            top = Math.Max(top, context.Y);
            OfficeImageProjection projection = CreateImageProjection(image, left, top, width, height);
            (double boundsLeft, double boundsTop, double boundsRight, double boundsBottom) = projection.GetDestinationBounds();
            if (boundsLeft < 0D || boundsTop < 0D || boundsRight > context.Drawing.Width || boundsBottom > context.Drawing.Height) {
                AddDiagnostic(
                    diagnostics,
                    "unsupported-word-floating-image",
                    "Skipped a Word image because its top-and-bottom anchor projects outside the current page preview.",
                    DescribeImage(image));
                return false;
            }

            double distanceFromBottom = GetAnchorDistancePoints(image._Image.Anchor?.DistanceFromBottom);
            if (!EnsureVerticalSpace(context, boundsBottom + distanceFromBottom - context.Y, diagnostics)) {
                return false;
            }

            context.Drawing.AddImage(bytes, image.ContentType, projection, DescribeImage(image));
            context.Y = boundsBottom + distanceFromBottom + ParagraphGapPoints;
            return true;
        }

        private static bool AddSideWrappedAnchoredImage(WordImage image, WrapTextImage wrapText, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics) {
            if (!TryReadEmbeddedImage(image, diagnostics, out byte[] bytes, out double width, out double height)) {
                return false;
            }

            FitImageToWidth(context.ContentWidth, ref width, ref height);
            if (!TryGetNoWrapAnchorPlacement(image, context, width, height, out double left, out double top)) {
                AddDiagnostic(
                    diagnostics,
                    "unsupported-word-floating-image",
                    "Skipped a Word image because its " + DescribeWrapText(wrapText) + " anchor position could not be resolved.",
                    DescribeImage(image));
                return false;
            }

            OfficeImageProjection projection = CreateImageProjection(image, left, top, width, height);
            (double boundsLeft, double boundsTop, double boundsRight, double boundsBottom) = projection.GetDestinationBounds();
            if (boundsLeft < 0D || boundsTop < 0D || boundsRight > context.Drawing.Width || boundsBottom > context.Drawing.Height) {
                AddDiagnostic(
                    diagnostics,
                    "unsupported-word-floating-image",
                    "Skipped a Word image because its " + DescribeWrapText(wrapText) + " anchor projects outside the current page preview.",
                    DescribeImage(image));
                return false;
            }

            context.Drawing.AddImage(bytes, image.ContentType, projection, DescribeImage(image));
            Anchor? anchor = image._Image.Anchor;
            context.AddTextExclusion(
                Math.Max(context.Left, boundsLeft - GetAnchorDistancePoints(anchor?.DistanceFromLeft)),
                Math.Max(0D, boundsTop - GetAnchorDistancePoints(anchor?.DistanceFromTop)),
                Math.Min(context.Left + context.ContentWidth, boundsRight + GetAnchorDistancePoints(anchor?.DistanceFromRight)),
                Math.Min(context.ContentBottom, boundsBottom + GetAnchorDistancePoints(anchor?.DistanceFromBottom)),
                GetTextWrapSide(anchor, wrapText));
            if (wrapText == WrapTextImage.Tight || wrapText == WrapTextImage.Through) {
                AddDiagnostic(
                    diagnostics,
                    "limited-word-floating-image-wrap",
                    "Rendered a Word " + DescribeWrapText(wrapText) + " image with a rectangular text exclusion because dependency-free export does not yet implement polygon or transparent-region wrapping.",
                    DescribeImage(image));
            }

            return true;
        }

        private static void FitImageToWidth(double maxWidth, ref double width, ref double height) {
            if (width <= 0D || height <= 0D || width <= maxWidth) {
                return;
            }

            double ratio = maxWidth / width;
            width = maxWidth;
            height *= ratio;
        }

        private static bool TryReadEmbeddedImage(WordImage image, List<OfficeImageExportDiagnostic> diagnostics, out byte[] bytes, out double width, out double height) {
            bytes = Array.Empty<byte>();
            width = 0D;
            height = 0D;

            if (image.IsExternal) {
                AddDiagnostic(diagnostics, "unsupported-word-external-image", "Skipped a Word image because dependency-free export does not fetch external image relationships.", DescribeImage(image));
                return false;
            }

            try {
                bytes = image.GetBytes();
            } catch (InvalidOperationException) {
                AddDiagnostic(diagnostics, "unsupported-word-image", "Skipped a Word image because its embedded bytes could not be read.", DescribeImage(image));
                return false;
            }

            if (bytes.Length == 0) {
                AddDiagnostic(diagnostics, "unsupported-word-image", "Skipped a Word image because its embedded bytes are empty.", DescribeImage(image));
                return false;
            }

            width = Helpers.ConvertPixelsToPoints(image.Width ?? 64D);
            height = Helpers.ConvertPixelsToPoints(image.Height ?? 64D);
            if (width <= 0D || height <= 0D) {
                AddDiagnostic(diagnostics, "unsupported-word-image", "Skipped a Word image because its dimensions could not be resolved.", DescribeImage(image));
                return false;
            }

            return true;
        }

        private static OfficeImageProjection CreateImageProjection(WordImage image, double left, double top, double width, double height) =>
            new OfficeImageProjection(
                new OfficeImagePlacement(left, top, width, height),
                CreateImageSourceCrop(image),
                image.Rotation ?? 0D,
                left + (width / 2D),
                top + (height / 2D),
                image.HorizontalFlip == true,
                image.VerticalFlip == true);

        private static OfficeImageSourceCrop CreateImageSourceCrop(WordImage image) =>
            OfficeImageSourceCrop.FromClampedFractions(
                ConvertWordCropFraction(image.CropLeft),
                ConvertWordCropFraction(image.CropTop),
                ConvertWordCropFraction(image.CropRight),
                ConvertWordCropFraction(image.CropBottom));

        private static double ConvertWordCropFraction(int? value) =>
            value.HasValue ? value.Value / 100000D : 0D;

        private static bool TryGetNoWrapAnchorPlacement(WordImage image, WordImageFlowContext context, double width, double height, out double left, out double top) {
            left = 0D;
            top = 0D;

            Anchor? anchor = image._Image.Anchor;
            if (anchor == null) {
                return false;
            }

            left = ResolveHorizontalAnchorPosition(anchor.HorizontalPosition, context, width);
            top = ResolveVerticalAnchorPosition(anchor.VerticalPosition, context, height);
            return IsFinite(left) && IsFinite(top);
        }

        private static bool IsNoWrapAnchoredImage(WrapTextImage wrapText) =>
            wrapText == WrapTextImage.BehindText || wrapText == WrapTextImage.InFrontOfText;

        private static bool IsSideWrappedAnchoredImage(WrapTextImage wrapText) =>
            wrapText == WrapTextImage.Square || wrapText == WrapTextImage.Tight || wrapText == WrapTextImage.Through;

        private static WordTextWrapSide GetTextWrapSide(Anchor? anchor, WrapTextImage wrapText) {
            WrapTextValues? wrapValue = null;
            if (wrapText == WrapTextImage.Square) {
                wrapValue = anchor?.Elements<WrapSquare>().FirstOrDefault()?.WrapText?.Value;
            } else if (wrapText == WrapTextImage.Tight) {
                wrapValue = anchor?.Elements<WrapTight>().FirstOrDefault()?.WrapText?.Value;
            } else if (wrapText == WrapTextImage.Through) {
                wrapValue = anchor?.Elements<WrapThrough>().FirstOrDefault()?.WrapText?.Value;
            }

            if (wrapValue == WrapTextValues.Left) {
                return WordTextWrapSide.Left;
            }

            if (wrapValue == WrapTextValues.Right) {
                return WordTextWrapSide.Right;
            }

            return WordTextWrapSide.Largest;
        }

        private static string DescribeWrapText(WrapTextImage wrapText) {
            switch (wrapText) {
                case WrapTextImage.Square:
                    return "square-wrap";
                case WrapTextImage.Tight:
                    return "tight-wrap";
                case WrapTextImage.Through:
                    return "through-wrap";
                default:
                    return "wrapped";
            }
        }

        private static double ResolveHorizontalAnchorPosition(HorizontalPosition? position, WordImageFlowContext context, double width) {
            double offset = GetPositionOffsetPoints(position?.PositionOffset);
            string? alignment = position?.HorizontalAlignment?.Text;
            if (!string.IsNullOrWhiteSpace(alignment)) {
                double containerLeft = IsPageRelative(position) ? 0D : context.Left;
                double containerWidth = IsPageRelative(position) ? context.Drawing.Width : context.ContentWidth;
                if (string.Equals(alignment, "center", StringComparison.OrdinalIgnoreCase)) {
                    return containerLeft + ((containerWidth - width) / 2D) + offset;
                }

                if (string.Equals(alignment, "right", StringComparison.OrdinalIgnoreCase)) {
                    return containerLeft + containerWidth - width + offset;
                }

                if (string.Equals(alignment, "left", StringComparison.OrdinalIgnoreCase)) {
                    return containerLeft + offset;
                }
            }

            return IsPageRelative(position) ? offset : context.Left + offset;
        }

        private static double ResolveVerticalAnchorPosition(VerticalPosition? position, WordImageFlowContext context, double height) {
            double offset = GetPositionOffsetPoints(position?.PositionOffset);
            string? alignment = position?.VerticalAlignment?.Text;
            if (!string.IsNullOrWhiteSpace(alignment)) {
                double containerTop = IsPageRelative(position) ? 0D : context.Y;
                double containerBottom = IsPageRelative(position) ? context.Drawing.Height : context.ContentBottom;
                double containerHeight = Math.Max(0D, containerBottom - containerTop);
                if (string.Equals(alignment, "center", StringComparison.OrdinalIgnoreCase)) {
                    return containerTop + ((containerHeight - height) / 2D) + offset;
                }

                if (string.Equals(alignment, "bottom", StringComparison.OrdinalIgnoreCase)) {
                    return containerBottom - height + offset;
                }

                if (string.Equals(alignment, "top", StringComparison.OrdinalIgnoreCase)) {
                    return containerTop + offset;
                }
            }

            return IsPageRelative(position) ? offset : context.Y + offset;
        }

        private static bool IsPageRelative(HorizontalPosition? position) =>
            position?.RelativeFrom?.Value == HorizontalRelativePositionValues.Page;

        private static bool IsPageRelative(VerticalPosition? position) =>
            position?.RelativeFrom?.Value == VerticalRelativePositionValues.Page;

        private static double GetPositionOffsetPoints(PositionOffset? positionOffset) {
            if (positionOffset == null || string.IsNullOrWhiteSpace(positionOffset.Text)) {
                return 0D;
            }

            return long.TryParse(positionOffset.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emus)
                ? Helpers.ConvertEmusToPoints(emus)
                : 0D;
        }

        private static double GetAnchorDistancePoints(DocumentFormat.OpenXml.UInt32Value? distance) =>
            distance?.Value != null ? Helpers.ConvertEmusToPoints(distance.Value) : 0D;

        private static bool IsFinite(double value) =>
            !double.IsNaN(value) && !double.IsInfinity(value);

        private static void AddSvgImageDiagnostics(OfficeDrawing drawing, List<OfficeImageExportDiagnostic> diagnostics) {
            foreach (OfficeDrawingImage image in drawing.Images) {
                byte[] bytes = image.Bytes;
                if (!OfficeSvgImageRenderer.TryCreateDataUri(image.ContentType, bytes, null, out _)) {
                    AddImageDiagnostic(
                        diagnostics,
                        "unsupported-word-image-svg",
                        "Skipped a Word image in SVG output because its content type is not embeddable as an SVG image element.",
                        image);
                }
            }
        }

        private static void AddRasterImageDiagnostics(OfficeDrawing drawing, List<OfficeImageExportDiagnostic> diagnostics) {
            foreach (OfficeDrawingImage image in drawing.Images) {
                if (!OfficeRasterImageDecoder.TryDecode(image.Bytes, out _)) {
                    AddImageDiagnostic(
                        diagnostics,
                        "unsupported-word-image-raster",
                        "Skipped a Word image in PNG output because dependency-free raster export currently decodes PNG and uncompressed BMP image bytes only.",
                        image);
                }
            }
        }

        private static void AddImageDiagnostic(List<OfficeImageExportDiagnostic> diagnostics, string code, string message, OfficeDrawingImage image) {
            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                code,
                message,
                string.IsNullOrWhiteSpace(image.AlternativeText) ? "Word image" : image.AlternativeText));
        }

        private static string DescribeImage(WordImage image) {
            if (!string.IsNullOrWhiteSpace(image.Description)) {
                return image.Description!;
            }

            if (!string.IsNullOrWhiteSpace(image.Title)) {
                return image.Title!;
            }

            if (!string.IsNullOrWhiteSpace(image.FileName)) {
                return image.FileName!;
            }

            return "Word image";
        }
    }
}
