using System.Collections.Generic;
using System.Globalization;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    internal static partial class PdfWordConverter {
        private static void AddImage(
            WordDocument document,
            PdfCore.PdfLogicalPage page,
            PdfCore.PdfLogicalImage image,
            PdfCore.PdfImagePlacement? placement,
            PdfToWordOptions options) {
            PdfCore.PdfImagePlacementImportAssessment assessment =
                PdfCore.PdfImagePlacementImportPolicy.Analyze(page, image, placement);
            if (assessment.IsSuppressed) {
                ReportSuppressedImagePlacement(image, assessment, options);
                if (assessment.Disposition == PdfCore.PdfImagePlacementImportDisposition.SuppressUnplaced &&
                    options.IncludeImagePlaceholders) {
                    AddImagePlaceholder(document, image, "no-visible-placement", options);
                }
                return;
            }
            if (!assessment.CanImport) {
                ReportUnsafeImagePlacement(image, assessment, options);
                if (options.IncludeImagePlaceholders) {
                    AddImagePlaceholder(document, image, "unsafe-placement-effects", options);
                }
                return;
            }

            if (options.ImportImages && TryAddEmbeddedImage(document, page, image, placement, assessment, options)) {
                return;
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
            }
        }

        private static bool TryAddEmbeddedImage(
            WordDocument document,
            PdfCore.PdfLogicalPage page,
            PdfCore.PdfLogicalImage image,
            PdfCore.PdfImagePlacement? placement,
            PdfCore.PdfImagePlacementImportAssessment assessment,
            PdfToWordOptions options) {
            PdfCore.PdfExtractedImage source = image.SourceImage;
            if (!source.IsImageFile || source.Bytes.Length == 0) {
                AddImageSkippedWarning(image, "PDF image stream is not exposed as a complete image file payload.");
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
                width = PdfPointsToWordPixels(placement.Width);
                height = PdfPointsToWordPixels(placement.Height);
            }

            try {
                using var stream = new MemoryStream(source.Bytes);
                WordParagraph imageParagraph = document.AddParagraph();
                ApplySourceParagraphSpacing(imageParagraph, options);
                string description = "Imported PDF image " + image.ResourceName + " from page " + image.PageNumber.ToString(CultureInfo.InvariantCulture);
                WordImage embeddedImage;
                if (options.PreserveImagePlacementPosition && placement != null && placement.IsAxisAligned && page.RotationDegrees == 0) {
                    PdfCore.PdfSelectionQuad visual = page.MapUserSpaceRectangleToVisual(
                        placement.X,
                        placement.Y,
                        placement.X + placement.Width,
                        placement.Y + placement.Height);
                    embeddedImage = imageParagraph.InsertImage(
                        stream,
                        fileName,
                        PdfPointsToWordPixels(visual.Width),
                        PdfPointsToWordPixels(visual.Height),
                        WordImageTextWrapping.InFrontOfText,
                        description);
                    embeddedImage.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
                    embeddedImage.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
                    embeddedImage.HorizontalPositionOffset = PdfPointsToEmu(visual.Left);
                    embeddedImage.VerticalPositionOffset = PdfPointsToEmu(visual.Top);
                } else {
                    embeddedImage = imageParagraph.InsertImage(stream, fileName, width, height, description: description);
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
            } catch (Exception ex) when (ex is ArgumentException || ex is InvalidOperationException || ex is NotSupportedException) {
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

        private static void ApplyImagePlacementEffects(
            WordImage embeddedImage,
            PdfCore.PdfLogicalImage image,
            PdfCore.PdfImagePlacementImportAssessment assessment,
            PdfToWordOptions options) {
            if (assessment.HasNonDefaultOpacity) {
                embeddedImage.Transparency = (int)Math.Round(
                    (1D - assessment.Opacity) * 100D,
                    MidpointRounding.AwayFromZero);
                AddWarning(
                    options,
                    "PdfImageOpacityMapped",
                    "Page " + image.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Image",
                    "PDF image opacity was mapped to native Word picture transparency.",
                    PdfCore.PdfConversionWarningSeverity.Information,
                    OfficeConversionLossKind.None,
                    new Dictionary<string, string> {
                        ["ResourceName"] = image.ResourceName,
                        ["Opacity"] = assessment.Opacity.ToString("R", CultureInfo.InvariantCulture)
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
            string code = assessment.Disposition == PdfCore.PdfImagePlacementImportDisposition.SuppressInvisible
                ? "PdfInvisibleImagePlacementSuppressed"
                : "PdfUnplacedImageResourceNotEmbedded";
            string message = assessment.Disposition == PdfCore.PdfImagePlacementImportDisposition.SuppressInvisible
                ? "A fully transparent PDF image placement was suppressed instead of exposing its raw image pixels."
                : "An extracted image resource without a visible page placement was not embedded as raw pixels.";
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
