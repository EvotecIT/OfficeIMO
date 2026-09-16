using System.Collections.Generic;
using System.Globalization;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    internal static partial class PdfWordConverter {
        private static bool ShouldQueueImage(
            PdfCore.PdfLogicalPage page,
            PdfCore.PdfLogicalImage image,
            PdfCore.PdfImagePlacement? placement,
            PdfToWordOptions options) {
            PdfCore.PdfImagePlacementImportAssessment assessment =
                PdfCore.PdfImagePlacementImportPolicy.Analyze(page, image, placement);
            if (assessment.IsSuppressed) {
                if (assessment.Disposition == PdfCore.PdfImagePlacementImportDisposition.SuppressUnplaced &&
                    options.IncludeImagePlaceholders) {
                    return true;
                }

                ReportSuppressedImagePlacement(image, assessment, options);
                return false;
            }

            if (!assessment.CanImport && !options.IncludeImagePlaceholders) {
                ReportUnsafeImagePlacement(image, assessment, options);
                return false;
            }

            return true;
        }

        private static bool AddImage(
            WordDocument document,
            PdfCore.PdfLogicalPage page,
            PdfCore.PdfLogicalImage image,
            PdfCore.PdfImagePlacement? placement,
            bool sourcePageSizeApplied,
            PdfToWordOptions options) {
            PdfCore.PdfImagePlacementImportAssessment assessment =
                PdfCore.PdfImagePlacementImportPolicy.Analyze(page, image, placement);
            if (assessment.IsSuppressed) {
                ReportSuppressedImagePlacement(image, assessment, options);
                if (assessment.Disposition == PdfCore.PdfImagePlacementImportDisposition.SuppressUnplaced &&
                    options.IncludeImagePlaceholders) {
                    AddImagePlaceholder(document, image, "no-visible-placement", options);
                    return true;
                }
                return false;
            }
            if (!assessment.CanImport) {
                ReportUnsafeImagePlacement(image, assessment, options);
                if (options.IncludeImagePlaceholders) {
                    AddImagePlaceholder(document, image, "unsafe-placement-effects", options);
                    return true;
                }
                return false;
            }

            if (options.ImportImages && TryAddEmbeddedImage(document, page, image, placement, assessment, sourcePageSizeApplied, options)) {
                return true;
            }

            if (options.IncludeImagePlaceholders) {
                if (!options.ImportImages) {
                    AddWarning(
                        options,
                        "PdfImagePlaceholder",
                        "Page " + image.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Image",
                        "PDF image pixels were not imported; an editable text placeholder represents the omitted image resource.",
                        PdfCore.PdfConversionWarningSeverity.Warning,
                        OfficeConversionLossKind.Omission,
                        new Dictionary<string, string> {
                            ["ResourceName"] = image.ResourceName,
                            ["MimeType"] = image.MimeType ?? string.Empty
                        });
                }
                AddImagePlaceholder(
                    document,
                    image,
                    options.ImportImages ? "unsupported-image-payload" : "image-import-disabled",
                    options);
                return true;
            }
            return false;
        }

        private static bool TryAddEmbeddedImage(
            WordDocument document,
            PdfCore.PdfLogicalPage page,
            PdfCore.PdfLogicalImage image,
            PdfCore.PdfImagePlacement? placement,
            PdfCore.PdfImagePlacementImportAssessment assessment,
            bool sourcePageSizeApplied,
            PdfToWordOptions options) {
            PdfCore.PdfExtractedImage source = image.SourceImage;
            if (!source.IsImageFile || source.Bytes.Length == 0) {
                AddImageSkippedWarning(image, "PDF image stream is not exposed as a complete image file payload.");
                return false;
            }
            if (string.Equals(source.MimeType, "image/j2c", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(source.FileExtension?.TrimStart('.'), "j2c", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(source.FileExtension?.TrimStart('.'), "j2k", StringComparison.OrdinalIgnoreCase)) {
                AddImageSkippedWarning(image, "Raw JPEG 2000 codestreams are not supported by the Word image-part format.");
                return false;
            }

            string extension = ResolveImageExtension(source);
            if (string.IsNullOrWhiteSpace(extension)) {
                AddImageSkippedWarning(image, "PDF image file extension could not be resolved for Word embedding.");
                return false;
            }

            string fileName = BuildImageFileName(image, extension);
            double? width = null;
            double? height = null;
            if (options.PreserveImagePlacementSize && placement != null && placement.Width > 0 && placement.Height > 0) {
                if (page.RotationDegrees != 0) {
                    double userUnit = page.UserUnit.GetValueOrDefault(1D);
                    width = PdfPointsToWordPixels(placement.Width * userUnit);
                    height = PdfPointsToWordPixels(placement.Height * userUnit);
                } else {
                    PdfCore.PdfSelectionQuad visualBounds = page.MapUserSpaceRectangleToVisual(
                        placement.X,
                        placement.Y,
                        placement.X + placement.Width,
                        placement.Y + placement.Height);
                    width = PdfPointsToWordPixels(visualBounds.Right - visualBounds.Left);
                    height = PdfPointsToWordPixels(visualBounds.Bottom - visualBounds.Top);
                }
            }

            try {
                using var stream = new MemoryStream(source.Bytes);
                WordParagraph imageParagraph = document.AddParagraph();
                ApplySourceParagraphSpacing(imageParagraph, options);
                string description = "Imported PDF image " + image.ResourceName + " from page " + image.PageNumber.ToString(CultureInfo.InvariantCulture);
                WordImage embeddedImage;
                if (sourcePageSizeApplied && options.PreserveImagePlacementPosition && placement != null && placement.IsAxisAligned && page.RotationDegrees == 0) {
                    imageParagraph.LineSpacingRule = WordLineSpacingRule.Exact;
                    imageParagraph.LineSpacing = 1;
                    PdfCore.PdfSelectionQuad visual = page.MapUserSpaceRectangleToVisual(
                        placement.X,
                        placement.Y,
                        placement.X + placement.Width,
                        placement.Y + placement.Height);
                    WordImageTextWrapping wrapping = HasOverlappingTextPaintedAfter(page, placement)
                        ? WordImageTextWrapping.BehindText
                        : WordImageTextWrapping.InFrontOfText;
                    embeddedImage = imageParagraph.InsertImage(
                        stream,
                        fileName,
                        width,
                        height,
                        wrapping,
                        description);
                    embeddedImage.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
                    embeddedImage.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
                    embeddedImage.HorizontalPositionOffset = PdfPointsToEmu(visual.Left);
                    embeddedImage.VerticalPositionOffset = PdfPointsToEmu(visual.Top);
                    embeddedImage.ZOrder = GetWordImageZOrder(page, placement);
                } else {
                    embeddedImage = imageParagraph.InsertImage(stream, fileName, width, height, description: description);
                    if (placement != null && page.RotationDegrees != 0) {
                        embeddedImage.Rotation = page.RotationDegrees;
                    }
                }
                if (placement?.IsAxisAligned == true) {
                    embeddedImage.HorizontalFlip = placement.A < 0D;
                    embeddedImage.VerticalFlip = placement.D < 0D;
                }
                ApplyImagePlacementEffects(embeddedImage, image, assessment, options);
                AddWarning(
                    options,
                    "PdfImageEmbedded",
                    "Page " + image.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Image",
                    "PDF image content was embedded as a native Word image.",
                    PdfCore.PdfConversionWarningSeverity.Information,
                    new Dictionary<string, string> {
                        ["ResourceName"] = image.ResourceName,
                        ["Width"] = image.Width.ToString(CultureInfo.InvariantCulture),
                        ["Height"] = image.Height.ToString(CultureInfo.InvariantCulture),
                        ["MimeType"] = image.MimeType ?? string.Empty,
                        ["PlacementWidth"] = placement?.Width.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                        ["PlacementHeight"] = placement?.Height.ToString(CultureInfo.InvariantCulture) ?? string.Empty
                    });
                return true;
            } catch (Exception ex) when (ex is ArgumentException || ex is InvalidDataException || ex is InvalidOperationException || ex is NotSupportedException) {
                AddImageSkippedWarning(image, "Word image embedding rejected the extracted PDF image payload: " + ex.Message);
                return false;
            }

            void AddImageSkippedWarning(PdfCore.PdfLogicalImage skippedImage, string message) {
                AddWarning(
                    options,
                    "PdfImageEmbeddingSkipped",
                    "Page " + skippedImage.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Image",
                    message,
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    OfficeConversionLossKind.Omission,
                    new Dictionary<string, string> {
                        ["ResourceName"] = skippedImage.ResourceName,
                        ["MimeType"] = skippedImage.MimeType ?? string.Empty,
                        ["IsImageFile"] = skippedImage.SourceImage.IsImageFile ? "true" : "false"
                    });
            }
        }

        private static uint GetWordImageZOrder(
            PdfCore.PdfLogicalPage page,
            PdfCore.PdfImagePlacement placement) {
            const uint baseZOrder = 251658240U;
            int rank = page.Images
                .SelectMany(static image => image.Placements)
                .OrderBy(static candidate => candidate.PaintOrder)
                .ThenBy(static candidate => candidate.ObjectNumber)
                .ThenBy(static candidate => candidate.ResourceName, StringComparer.Ordinal)
                .TakeWhile(candidate => !ReferenceEquals(candidate, placement))
                .Count();
            return baseZOrder + (uint)Math.Min(rank, (long)uint.MaxValue - baseZOrder);
        }

        private static bool HasOverlappingTextPaintedAfter(
            PdfCore.PdfLogicalPage page,
            PdfCore.PdfImagePlacement placement) {
            double imageLeft = placement.X;
            double imageRight = placement.X + placement.Width;
            double imageBottom = placement.Y;
            double imageTop = placement.Y + placement.Height;
            for (int blockIndex = 0; blockIndex < page.TextBlocks.Count; blockIndex++) {
                IReadOnlyList<PdfCore.PdfTextSpan> spans = page.TextBlocks[blockIndex].Spans;
                for (int spanIndex = 0; spanIndex < spans.Count; spanIndex++) {
                    PdfCore.PdfTextSpan span = spans[spanIndex];
                    if (!span.IsVisible || span.PaintOrder <= placement.PaintOrder || string.IsNullOrEmpty(span.Text)) {
                        continue;
                    }

                    double radians = span.RotationDegrees * Math.PI / 180D;
                    double advance = span.Advance;
                    if (advance == 0D) {
                        advance = Math.Max(1D, span.FontSize * 0.55D * span.Text.Length);
                    }
                    double directionX = Math.Cos(radians);
                    double directionY = Math.Sin(radians);
                    double normalX = -directionY;
                    double normalY = directionX;
                    double endX = span.X + advance * directionX;
                    double endY = span.Y + advance * directionY;
                    double ascent = Math.Max(1D, span.FontSize * 0.8D);
                    double descent = Math.Max(0.25D, span.FontSize * 0.2D);
                    double startTopX = span.X + normalX * ascent;
                    double startTopY = span.Y + normalY * ascent;
                    double startBottomX = span.X - normalX * descent;
                    double startBottomY = span.Y - normalY * descent;
                    double endTopX = endX + normalX * ascent;
                    double endTopY = endY + normalY * ascent;
                    double endBottomX = endX - normalX * descent;
                    double endBottomY = endY - normalY * descent;
                    double textLeft = Math.Min(Math.Min(startTopX, startBottomX), Math.Min(endTopX, endBottomX));
                    double textRight = Math.Max(Math.Max(startTopX, startBottomX), Math.Max(endTopX, endBottomX));
                    double textBottom = Math.Min(Math.Min(startTopY, startBottomY), Math.Min(endTopY, endBottomY));
                    double textTop = Math.Max(Math.Max(startTopY, startBottomY), Math.Max(endTopY, endBottomY));
                    if (textRight > imageLeft && textLeft < imageRight &&
                        textTop > imageBottom && textBottom < imageTop) {
                        return true;
                    }
                }
            }
            return false;
        }

        private static void ApplyImagePlacementEffects(
            WordImage embeddedImage,
            PdfCore.PdfLogicalImage image,
            PdfCore.PdfImagePlacementImportAssessment assessment,
            PdfToWordOptions options) {
            if (assessment.HasNonDefaultOpacity) {
                embeddedImage.Transparency = assessment.MappedTransparencyPercent;
                bool opacityIsOmitted = assessment.MappedOpacityIsOmitted;
                AddWarning(
                    options,
                    "PdfImageOpacityMapped",
                    "Page " + image.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Image",
                    opacityIsOmitted
                        ? "PDF image opacity was below Word picture-transparency precision, so the image was made fully transparent."
                        : "PDF image opacity was mapped to native Word picture transparency.",
                    opacityIsOmitted
                        ? PdfCore.PdfConversionWarningSeverity.Warning
                        : PdfCore.PdfConversionWarningSeverity.Information,
                    opacityIsOmitted ? OfficeConversionLossKind.Omission : OfficeConversionLossKind.None,
                    new Dictionary<string, string> {
                        ["ResourceName"] = image.ResourceName,
                        ["Opacity"] = assessment.Opacity.ToString("R", CultureInfo.InvariantCulture),
                        ["MappedTransparencyPercent"] = assessment.MappedTransparencyPercent.ToString(CultureInfo.InvariantCulture)
                    });
            }
            if (assessment.HasNonNormalBlendMode) {
                AddWarning(
                    options,
                    "PdfImageBlendModeApproximated",
                    "Page " + image.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Image",
                    "Word does not reproduce the PDF image blend mode; the safely visible image pixels were embedded with normal compositing.",
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    OfficeConversionLossKind.Approximation,
                    new Dictionary<string, string> {
                        ["ResourceName"] = image.ResourceName,
                        ["BlendMode"] = assessment.BlendMode.ToString()
                    });
            }
        }

        private static void ReportSuppressedImagePlacement(
            PdfCore.PdfLogicalImage image,
            PdfCore.PdfImagePlacementImportAssessment assessment,
            PdfToWordOptions options) {
            string code = assessment.Disposition switch {
                PdfCore.PdfImagePlacementImportDisposition.SuppressInvisible => "PdfInvisibleImagePlacementSuppressed",
                PdfCore.PdfImagePlacementImportDisposition.SuppressOutsideVisibleArea => "PdfNonVisibleImagePlacementSuppressed",
                _ => "PdfUnplacedImageResourceNotEmbedded"
            };
            string message = assessment.Disposition switch {
                PdfCore.PdfImagePlacementImportDisposition.SuppressInvisible =>
                    "A fully transparent PDF image placement was suppressed instead of exposing its raw image pixels.",
                PdfCore.PdfImagePlacementImportDisposition.SuppressOutsideVisibleArea =>
                    "A PDF image placement with no visible page intersection was suppressed instead of exposing its raw image pixels.",
                _ => "An extracted image resource without a visible page placement was not embedded as raw pixels."
            };
            AddWarning(
                options,
                code,
                "Page " + image.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Image",
                message,
                PdfCore.PdfConversionWarningSeverity.Information,
                OfficeConversionLossKind.None,
                new Dictionary<string, string> { ["ResourceName"] = image.ResourceName });
        }

        private static void ReportUnsafeImagePlacement(
            PdfCore.PdfLogicalImage image,
            PdfCore.PdfImagePlacementImportAssessment assessment,
            PdfToWordOptions options) {
            (string code, string message) = assessment.Disposition switch {
                PdfCore.PdfImagePlacementImportDisposition.OmitClippedPixels => (
                    "PdfImageClipNotSafelyEditable",
                    "The raw PDF image was not embedded because its clip hides source pixels that an editable Word picture could reveal."),
                PdfCore.PdfImagePlacementImportDisposition.OmitSoftMask => (
                    "PdfImagePlacementSoftMaskNotSafelyEditable",
                    "The raw PDF image was not embedded because its placement soft mask cannot be reproduced safely as an editable Word picture."),
                PdfCore.PdfImagePlacementImportDisposition.OmitUnsupportedBlendMode => (
                    "PdfImageUnsupportedBlendModeNotSafelyEditable",
                    "The raw PDF image was not embedded because its unsupported PDF blend mode cannot be reproduced safely as an editable Word picture."),
                PdfCore.PdfImagePlacementImportDisposition.OmitUnsupportedPaintEffect => (
                    "PdfImagePaintEffectNotSafelyEditable",
                    "The raw PDF image was not embedded because its PDF paint effect cannot be reproduced safely as an editable Word picture."),
                PdfCore.PdfImagePlacementImportDisposition.OmitUnsupportedTransform => (
                    "PdfImageTransformNotSafelyEditable",
                    "The raw PDF image was not embedded because its rotation or shear cannot be reproduced safely as an editable Word picture."),
                PdfCore.PdfImagePlacementImportDisposition.OmitUnappliedDecode => (
                    "PdfImageDecodeNotSafelyEditable",
                    "The raw PDF image was not embedded because its PDF decode mapping is not represented by the extracted JPEG 2000 payload."),
                _ => (
                    "PdfImageTransparencyMaskNotResolved",
                    "The raw PDF image was not embedded because its unresolved transparency mask could hide source pixels that an editable Word picture would reveal.")
            };
            AddWarning(
                options,
                code,
                "Page " + image.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Image",
                message,
                PdfCore.PdfConversionWarningSeverity.Warning,
                OfficeConversionLossKind.Omission,
                new Dictionary<string, string> {
                    ["ResourceName"] = image.ResourceName,
                    ["MimeType"] = image.MimeType ?? string.Empty,
                    ["MaskKind"] = image.SourceImage.TransparencyMaskKind ?? string.Empty
                });
        }

        private static void AddImagePlaceholder(
            WordDocument document,
            PdfCore.PdfLogicalImage image,
            string reason,
            PdfToWordOptions options) {
            string text = "[PDF image: page "
                + image.PageNumber.ToString(CultureInfo.InvariantCulture)
                + ", resource "
                + image.ResourceName
                + ", "
                + image.Width.ToString(CultureInfo.InvariantCulture)
                + "x"
                + image.Height.ToString(CultureInfo.InvariantCulture)
                + (image.MimeType == null ? string.Empty : ", " + image.MimeType)
                + ", "
                + reason
                + "]";
            WordParagraph paragraph = document.AddParagraph(text).SetItalic();
            ApplySourceParagraphSpacing(paragraph, options);
        }

        private static string ResolveImageExtension(PdfCore.PdfExtractedImage image) {
            if (!string.IsNullOrWhiteSpace(image.FileExtension)) {
                return image.FileExtension!.TrimStart('.');
            }

            switch (image.MimeType?.ToLowerInvariant()) {
                case "image/jpeg":
                    return "jpg";
                case "image/png":
                    return "png";
                case "image/gif":
                    return "gif";
                case "image/bmp":
                    return "bmp";
                case "image/tiff":
                    return "tif";
                default:
                    return string.Empty;
            }
        }

        private static string BuildImageFileName(PdfCore.PdfLogicalImage image, string extension) {
            string resourceName = string.IsNullOrWhiteSpace(image.ResourceName) ? "image" : image.ResourceName;
            var safe = new char[resourceName.Length];
            for (int i = 0; i < resourceName.Length; i++) {
                char ch = resourceName[i];
                safe[i] = char.IsLetterOrDigit(ch) || ch == '-' || ch == '_' ? ch : '_';
            }

            return "pdf-page-"
                + image.PageNumber.ToString(CultureInfo.InvariantCulture)
                + "-"
                + new string(safe)
                + "."
                + extension;
        }

        private static double PdfPointsToWordPixels(double points) => points * 96D / 72D;

        private static long PdfPointsToEmu(double points) => checked((long)Math.Round(points * 12_700D));
    }
}
