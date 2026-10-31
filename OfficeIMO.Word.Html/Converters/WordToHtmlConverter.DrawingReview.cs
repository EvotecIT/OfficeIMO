using AngleSharp.Dom;
using OfficeIMO.Html;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static void ApplyImageReviewMetadata(IElement element, WordImage image, WordHtmlExportOptions options) {
            if (!options.IncludeDrawingReviewMetadata) return;

            WordImageTextWrapping? wrap;
            try {
                wrap = image.WrapText;
            } catch (InvalidOperationException) {
                AddExportDiagnostic(options, "floating-image-metadata-simplified",
                    "A malformed Word image anchor could not expose deterministic floating-layout metadata.",
                    OfficeConversionLossKind.Approximation);
                return;
            }

            SetOutputAttribute(element, "data-officeimo-wrap", (wrap ?? WordImageTextWrapping.InLineWithText).ToString(), "ImageReview:wrap");
            SetOutputAttribute(element, "data-officeimo-anchor",
                wrap == WordImageTextWrapping.InLineWithText ? "inline" : "floating", "ImageReview:anchor");
            AppendOptionalAttribute(element, "data-officeimo-rotation", image.Rotation, "ImageReview:rotation");
            AppendOptionalAttribute(element, "data-officeimo-crop-top", (image.CropTopRatio is double cropRatio2 ? (int?)System.Math.Round(cropRatio2 * 100000D, System.MidpointRounding.AwayFromZero) : null), "ImageReview:crop-top");
            AppendOptionalAttribute(element, "data-officeimo-crop-right", (image.CropRightRatio is double cropRatio4 ? (int?)System.Math.Round(cropRatio4 * 100000D, System.MidpointRounding.AwayFromZero) : null), "ImageReview:crop-right");
            AppendOptionalAttribute(element, "data-officeimo-crop-bottom", (image.CropBottomRatio is double cropRatio6 ? (int?)System.Math.Round(cropRatio6 * 100000D, System.MidpointRounding.AwayFromZero) : null), "ImageReview:crop-bottom");
            AppendOptionalAttribute(element, "data-officeimo-crop-left", (image.CropLeftRatio is double cropRatio8 ? (int?)System.Math.Round(cropRatio8 * 100000D, System.MidpointRounding.AwayFromZero) : null), "ImageReview:crop-left");
            AppendOptionalAttribute(element, "data-officeimo-opacity", image.FixedOpacity, "ImageReview:opacity");
            AppendOptionalAttribute(element, "data-officeimo-transparency", image.Transparency, "ImageReview:transparency");
            AppendOptionalAttribute(element, "data-officeimo-brightness", image.LuminanceBrightness, "ImageReview:brightness");
            AppendOptionalAttribute(element, "data-officeimo-contrast", image.LuminanceContrast, "ImageReview:contrast");
            AppendOptionalAttribute(element, "data-officeimo-blur-radius", (image.BlurRadius?.ToEmus()), "ImageReview:blur-radius");
            AppendOptionalAttribute(element, "data-officeimo-grayscale", image.GrayScale, "ImageReview:grayscale");
            AppendOptionalAttribute(element, "data-officeimo-flip-horizontal", image.HorizontalFlip, "ImageReview:flip-horizontal");
            AppendOptionalAttribute(element, "data-officeimo-flip-vertical", image.VerticalFlip, "ImageReview:flip-vertical");

            if (wrap == WordImageTextWrapping.InLineWithText) return;
            try {
                SetOutputAttribute(element, "data-officeimo-horizontal-relative", image.HorizontalPositionRelativeFrom.ToString(), "ImageReview:horizontal-relative");
                SetOutputAttribute(element, "data-officeimo-vertical-relative", image.VerticalPositionRelativeFrom.ToString(), "ImageReview:vertical-relative");
                AppendOptionalAttribute(element, "data-officeimo-horizontal-offset-emu", (image.HorizontalPositionOffset?.ToEmus()), "ImageReview:horizontal-offset");
                AppendOptionalAttribute(element, "data-officeimo-vertical-offset-emu", (image.VerticalPositionOffset?.ToEmus()), "ImageReview:vertical-offset");
            } catch (InvalidOperationException) {
                AddExportDiagnostic(options, "floating-image-metadata-simplified",
                    "A floating Word image preserved its wrap policy, but malformed position metadata was omitted.",
                    OfficeConversionLossKind.Approximation);
            }

            AddExportDiagnostic(options, "floating-image-projected-for-review",
                "A floating Word image was exported with inert wrap and anchor metadata; browser layout does not reproduce Word's floating-object algorithm.",
                OfficeConversionLossKind.Approximation);
        }

        private static void AppendOptionalAttribute(IElement element, string name, int? value, string source) {
            if (value.HasValue) SetOutputAttribute(element, name, value.Value.ToString(CultureInfo.InvariantCulture), source);
        }

        private static void AppendOptionalAttribute(IElement element, string name, long? value, string source) {
            if (value.HasValue) SetOutputAttribute(element, name, value.Value.ToString(CultureInfo.InvariantCulture), source);
        }

        private static void AppendOptionalAttribute(IElement element, string name, bool? value, string source) {
            if (value.HasValue) SetOutputAttribute(element, name, value.Value ? "true" : "false", source);
        }
    }
}
