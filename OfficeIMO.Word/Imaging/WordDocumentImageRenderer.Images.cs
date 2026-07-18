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
                    if (!context.IsTargetPage) {
                        return false;
                    }

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
            double requiredHeight = Math.Max(height, boundsBottom - context.Y);
            if (!EnsureVerticalSpace(context, requiredHeight, diagnostics)) {
                return false;
            }

            if (Math.Abs(top - context.Y) > 0.000001D) {
                top = context.Y;
                projection = CreateImageProjection(image, left, top, width, height);
                (boundsLeft, boundsTop, boundsRight, boundsBottom) = projection.GetDestinationBounds();
            }

            if (boundsLeft < 0D || boundsTop < 0D || boundsRight > context.Drawing.Width || boundsBottom > context.Drawing.Height) {
                AddDiagnostic(
                    diagnostics,
                    "unsupported-word-image",
                    "Skipped a Word image because its inline projection extends outside the current page preview.",
                    DescribeImage(image));
                return false;
            }

            if (context.IsTargetPage) {
                context.Drawing.AddImage(bytes, image.ContentType, projection, DescribeImage(image));
            }

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
            double distanceFromBottom = GetAnchorDistancePoints(image._Image.Anchor?.DistanceFromBottom);
            OfficeImageProjection projection = CreateImageProjection(image, left, top, width, height);
            (double _, double _, double _, double projectedBottom) = projection.GetDestinationBounds();
            double requiredHeight = Math.Max(top + height, projectedBottom) + distanceFromBottom - context.Y;
            if (!EnsureVerticalSpace(context, requiredHeight, diagnostics)) {
                return false;
            }

            if (!TryGetNoWrapAnchorPlacement(image, context, width, height, out left, out top)) {
                AddDiagnostic(
                    diagnostics,
                    "unsupported-word-floating-image",
                    "Skipped a Word image because its top-and-bottom anchor position could not be resolved.",
                    DescribeImage(image));
                return false;
            }

            top = Math.Max(top, context.Y);
            projection = CreateImageProjection(image, left, top, width, height);
            (double boundsLeft, double boundsTop, double boundsRight, double boundsBottom) = projection.GetDestinationBounds();
            if (boundsLeft < 0D || boundsTop < 0D || boundsRight > context.Drawing.Width || boundsBottom > context.Drawing.Height) {
                AddDiagnostic(
                    diagnostics,
                    "unsupported-word-floating-image",
                    "Skipped a Word image because its top-and-bottom anchor projects outside the current page preview.",
                    DescribeImage(image));
                return false;
            }

            if (context.IsTargetPage) {
                context.Drawing.AddImage(bytes, image.ContentType, projection, DescribeImage(image));
            }

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

            if (context.IsTargetPage) {
                context.Drawing.AddImage(bytes, image.ContentType, projection, DescribeImage(image));
            }

            Anchor? anchor = image._Image.Anchor;
            double exclusionLeft = Math.Max(context.Left, boundsLeft - GetAnchorDistancePoints(anchor?.DistanceFromLeft));
            double exclusionTop = Math.Max(0D, boundsTop - GetAnchorDistancePoints(anchor?.DistanceFromTop));
            double exclusionRight = Math.Min(context.Left + context.ContentWidth, boundsRight + GetAnchorDistancePoints(anchor?.DistanceFromRight));
            double exclusionBottom = Math.Min(context.ContentBottom, boundsBottom + GetAnchorDistancePoints(anchor?.DistanceFromBottom));
            WordTextWrapSide wrapSide = GetTextWrapSide(anchor, wrapText);
            IReadOnlyList<OfficePoint> polygon = Array.Empty<OfficePoint>();
            bool usedAuthoredPolygon = (wrapText == WrapTextImage.Tight || wrapText == WrapTextImage.Through) &&
                TryCreateAuthoredWrapPolygonTextExclusion(anchor, exclusionLeft, exclusionTop, exclusionRight, exclusionBottom, out polygon);
            bool usedTransparentPolygon = false;
            if (usedAuthoredPolygon) {
                context.AddTextExclusion(polygon, wrapSide);
            } else if ((wrapText == WrapTextImage.Tight || wrapText == WrapTextImage.Through) &&
                TryCreateTransparentImageWrapPolygon(bytes, projection, out IReadOnlyList<OfficePoint> transparentPolygon)) {
                context.AddTextExclusion(transparentPolygon, wrapSide);
                usedTransparentPolygon = true;
            } else {
                context.AddTextExclusion(exclusionLeft, exclusionTop, exclusionRight, exclusionBottom, wrapSide);
            }

            if (context.IsTargetPage && (wrapText == WrapTextImage.Tight || wrapText == WrapTextImage.Through) && !usedAuthoredPolygon && !usedTransparentPolygon) {
                AddDiagnostic(
                    diagnostics,
                    "limited-word-floating-image-wrap",
                    "Rendered a Word " + DescribeWrapText(wrapText) + " image with a rectangular text exclusion because dependency-free export does not yet implement polygon or transparent-region wrapping.",
                    DescribeImage(image));
            }

            AdvanceFlowToAnchoredWrapTop(context, boundsTop);
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
                bytes = image.ToBytes();
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
            (double ContainerLeft, double ContainerWidth) container = ResolveHorizontalAnchorContainer(position, context);
            if (!string.IsNullOrWhiteSpace(alignment)) {
                if (string.Equals(alignment, "center", StringComparison.OrdinalIgnoreCase)) {
                    return container.ContainerLeft + ((container.ContainerWidth - width) / 2D) + offset;
                }

                if (string.Equals(alignment, "right", StringComparison.OrdinalIgnoreCase)) {
                    return container.ContainerLeft + container.ContainerWidth - width + offset;
                }

                if (string.Equals(alignment, "left", StringComparison.OrdinalIgnoreCase)) {
                    return container.ContainerLeft + offset;
                }
            }

            return container.ContainerLeft + offset;
        }

        private static double ResolveVerticalAnchorPosition(VerticalPosition? position, WordImageFlowContext context, double height) {
            double offset = GetPositionOffsetPoints(position?.PositionOffset);
            string? alignment = position?.VerticalAlignment?.Text;
            (double ContainerTop, double ContainerHeight) container = ResolveVerticalAnchorContainer(position, context);
            if (!string.IsNullOrWhiteSpace(alignment)) {
                if (string.Equals(alignment, "center", StringComparison.OrdinalIgnoreCase)) {
                    return container.ContainerTop + ((container.ContainerHeight - height) / 2D) + offset;
                }

                if (string.Equals(alignment, "bottom", StringComparison.OrdinalIgnoreCase)) {
                    return container.ContainerTop + container.ContainerHeight - height + offset;
                }

                if (string.Equals(alignment, "top", StringComparison.OrdinalIgnoreCase)) {
                    return container.ContainerTop + offset;
                }
            }

            return container.ContainerTop + offset;
        }

        private static (double ContainerLeft, double ContainerWidth) ResolveHorizontalAnchorContainer(HorizontalPosition? position, WordImageFlowContext context) {
            HorizontalRelativePositionValues? relativeFrom = position?.RelativeFrom?.Value;
            if (relativeFrom == HorizontalRelativePositionValues.Page) {
                return (0D, context.Drawing.Width);
            }

            if (relativeFrom == HorizontalRelativePositionValues.Margin) {
                return (context.Left, context.ContentWidth);
            }

            if (relativeFrom == HorizontalRelativePositionValues.LeftMargin) {
                return (0D, context.Left);
            }

            if (relativeFrom == HorizontalRelativePositionValues.RightMargin) {
                double left = context.Left + context.ContentWidth;
                return (left, Math.Max(0D, context.Drawing.Width - left));
            }

            if (relativeFrom == HorizontalRelativePositionValues.InsideMargin) {
                return IsOddWordPageIndex(context.PageIndex)
                    ? (0D, context.Left)
                    : GetRightMarginContainer(context);
            }

            if (relativeFrom == HorizontalRelativePositionValues.OutsideMargin) {
                return IsOddWordPageIndex(context.PageIndex)
                    ? GetRightMarginContainer(context)
                    : (0D, context.Left);
            }

            return (context.Left, context.ContentWidth);
        }

        private static (double ContainerLeft, double ContainerWidth) GetRightMarginContainer(WordImageFlowContext context) {
            double left = context.Left + context.ContentWidth;
            return (left, Math.Max(0D, context.Drawing.Width - left));
        }

        private static (double ContainerTop, double ContainerHeight) ResolveVerticalAnchorContainer(VerticalPosition? position, WordImageFlowContext context) {
            VerticalRelativePositionValues? relativeFrom = position?.RelativeFrom?.Value;
            if (relativeFrom == VerticalRelativePositionValues.Page) {
                return (0D, context.Drawing.Height);
            }

            if (relativeFrom == VerticalRelativePositionValues.Margin) {
                return (context.Top, Math.Max(0D, context.ContentBottom - context.Top));
            }

            if (relativeFrom == VerticalRelativePositionValues.TopMargin) {
                return (0D, context.Top);
            }

            if (relativeFrom == VerticalRelativePositionValues.BottomMargin) {
                return (context.ContentBottom, Math.Max(0D, context.Drawing.Height - context.ContentBottom));
            }

            return (context.Y, Math.Max(0D, context.ContentBottom - context.Y));
        }

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
