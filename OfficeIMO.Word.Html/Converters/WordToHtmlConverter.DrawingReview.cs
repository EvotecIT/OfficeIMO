using AngleSharp.Dom;
using OfficeIMO.Html;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static void ApplyImageReviewMetadata(IElement element, WordImage image, WordToHtmlOptions options) {
            if (!options.IncludeDrawingReviewMetadata) return;

            WordImageTextWrapping? wrap;
            try {
                wrap = image.WrapText;
            } catch (InvalidOperationException) {
                AddExportDiagnostic(options, "FloatingImageMetadataSimplified",
                    "A malformed Word image anchor could not expose deterministic floating-layout metadata.",
                    OfficeConversionLossKind.Approximation);
                return;
            }

            SetOutputAttribute(element, "data-officeimo-wrap", (wrap ?? WordImageTextWrapping.InLineWithText).ToString(), "ImageReview:wrap");
            SetOutputAttribute(element, "data-officeimo-anchor",
                wrap == WordImageTextWrapping.InLineWithText ? "inline" : "floating", "ImageReview:anchor");
            AppendOptionalAttribute(element, "data-officeimo-rotation", image.Rotation, "ImageReview:rotation");
            AppendOptionalAttribute(element, "data-officeimo-crop-top", image.CropTop, "ImageReview:crop-top");
            AppendOptionalAttribute(element, "data-officeimo-crop-right", image.CropRight, "ImageReview:crop-right");
            AppendOptionalAttribute(element, "data-officeimo-crop-bottom", image.CropBottom, "ImageReview:crop-bottom");
            AppendOptionalAttribute(element, "data-officeimo-crop-left", image.CropLeft, "ImageReview:crop-left");
            AppendOptionalAttribute(element, "data-officeimo-opacity", image.FixedOpacity, "ImageReview:opacity");
            AppendOptionalAttribute(element, "data-officeimo-transparency", image.Transparency, "ImageReview:transparency");
            AppendOptionalAttribute(element, "data-officeimo-brightness", image.LuminanceBrightness, "ImageReview:brightness");
            AppendOptionalAttribute(element, "data-officeimo-contrast", image.LuminanceContrast, "ImageReview:contrast");
            AppendOptionalAttribute(element, "data-officeimo-blur-radius", image.BlurRadius, "ImageReview:blur-radius");
            AppendOptionalAttribute(element, "data-officeimo-grayscale", image.GrayScale, "ImageReview:grayscale");
            AppendOptionalAttribute(element, "data-officeimo-flip-horizontal", image.HorizontalFlip, "ImageReview:flip-horizontal");
            AppendOptionalAttribute(element, "data-officeimo-flip-vertical", image.VerticalFlip, "ImageReview:flip-vertical");

            if (wrap == WordImageTextWrapping.InLineWithText) return;
            try {
                SetOutputAttribute(element, "data-officeimo-horizontal-relative", image.HorizontalPositionRelativeFrom.ToString(), "ImageReview:horizontal-relative");
                SetOutputAttribute(element, "data-officeimo-vertical-relative", image.VerticalPositionRelativeFrom.ToString(), "ImageReview:vertical-relative");
                AppendOptionalAttribute(element, "data-officeimo-horizontal-offset-emu", image.HorizontalPositionOffset, "ImageReview:horizontal-offset");
                AppendOptionalAttribute(element, "data-officeimo-vertical-offset-emu", image.VerticalPositionOffset, "ImageReview:vertical-offset");
            } catch (InvalidOperationException) {
                AddExportDiagnostic(options, "FloatingImageMetadataSimplified",
                    "A floating Word image preserved its wrap policy, but malformed position metadata was omitted.",
                    OfficeConversionLossKind.Approximation);
            }

            AddExportDiagnostic(options, "FloatingImageProjectedForReview",
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
