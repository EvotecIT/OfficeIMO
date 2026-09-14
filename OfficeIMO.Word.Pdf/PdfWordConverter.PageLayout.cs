using System.Collections.Generic;
using System.Globalization;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    internal static partial class PdfWordConverter {
        private static void ConfigureEditablePageSection(
            WordSection section,
            PdfCore.PdfLogicalPage page,
            PdfToWordOptions options) {
            (double pageWidth, double pageHeight) = GetVisualPageSize(page);
            if (pageWidth <= 0D || pageHeight <= 0D || pageWidth > 1584D || pageHeight > 1584D) {
                AddWarning(
                    options,
                    "PdfSourcePageSizeNotApplied",
                    "Page " + page.PageNumber.ToString(CultureInfo.InvariantCulture),
                    "The source PDF page size is outside Word's supported physical page-size range and was not applied.",
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    new Dictionary<string, string> {
                        ["WidthPoints"] = pageWidth.ToString(CultureInfo.InvariantCulture),
                        ["HeightPoints"] = pageHeight.ToString(CultureInfo.InvariantCulture)
                    });
                return;
            }

            if (double.IsNaN(options.EditablePageMarginPoints) ||
                double.IsInfinity(options.EditablePageMarginPoints) ||
                options.EditablePageMarginPoints < 0D) {
                throw new ArgumentOutOfRangeException(
                    nameof(options.EditablePageMarginPoints),
                    options.EditablePageMarginPoints,
                    "Editable page margin must be finite and nonnegative.");
            }

            double maximumMargin = Math.Max(0D, (Math.Min(pageWidth, pageHeight) - 1D) / 2D);
            double margin = Math.Min(options.EditablePageMarginPoints, maximumMargin);
            section.PageSettings.Orientation = pageWidth > pageHeight
                ? OfficePageOrientation.Landscape
                : OfficePageOrientation.Portrait;
            section.PageSettings.Width = checked((uint)Math.Round(pageWidth * 20D));
            section.PageSettings.Height = checked((uint)Math.Round(pageHeight * 20D));
            uint marginTwips = checked((uint)Math.Round(margin * 20D));
            section.Margins.Left = marginTwips;
            section.Margins.Right = marginTwips;
            section.Margins.Top = checked((int)marginTwips);
            section.Margins.Bottom = checked((int)marginTwips);
        }

        private static (double Width, double Height) GetVisualPageSize(PdfCore.PdfLogicalPage page) {
            PdfCore.PdfPageBox? box = page.Geometry.EffectiveBox;
            double scale = page.UserUnit.GetValueOrDefault(1D);
            double width = (box?.Width ?? page.Width) * scale;
            double height = (box?.Height ?? page.Height) * scale;
            return page.RotationDegrees is 90 or 270
                ? (height, width)
                : (width, height);
        }

        private static double GetEditableTypographyScale(
            PdfCore.PdfLogicalPage page,
            PdfToWordOptions options) {
            if (!options.PreserveSourcePageSize) return 1D;

            (double pageWidth, double pageHeight) = GetVisualPageSize(page);
            if (pageWidth <= 0D || pageHeight <= 0D || pageWidth > 1584D || pageHeight > 1584D) {
                return 1D;
            }

            double scale = page.UserUnit.GetValueOrDefault(1D);
            return scale > 0D && !double.IsNaN(scale) && !double.IsInfinity(scale) ? scale : 1D;
        }
    }
}
